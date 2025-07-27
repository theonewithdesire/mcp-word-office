"""
Recovery and graceful degradation utilities for Word MCP Server.
"""

import asyncio
import logging
import time
from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable, Dict, List, Optional

from .errors import (
    ConnectionError,
    DocumentError,
    ErrorCode,
    OperationError,
    WordMCPError,
)
from .logging import get_logger


class RecoveryStrategy(Enum):
    """Available recovery strategies."""

    RETRY = "retry"
    FALLBACK = "fallback"
    DEGRADE = "degrade"
    ABORT = "abort"


@dataclass
class RecoveryConfig:
    """Configuration for recovery attempts."""

    max_retries: int = 3
    retry_delay: float = 1.0
    backoff_multiplier: float = 2.0
    max_delay: float = 30.0
    timeout: float = 60.0


class GracefulDegradation:
    """Handles graceful degradation of functionality when operations fail."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or get_logger("recovery")
        self.degraded_features = set()
        self.fallback_handlers = {}
        self.recovery_strategies = {}

    def register_fallback(self, operation: str, fallback_handler: Callable):
        """Register a fallback handler for an operation."""
        self.fallback_handlers[operation] = fallback_handler
        self.logger.debug(f"Registered fallback handler for operation: {operation}")

    def register_recovery_strategy(
        self,
        error_code: str,
        strategy: RecoveryStrategy,
        handler: Optional[Callable] = None,
    ):
        """Register a recovery strategy for a specific error code."""
        self.recovery_strategies[error_code] = {
            "strategy": strategy,
            "handler": handler,
        }
        self.logger.debug(
            f"Registered recovery strategy {strategy.value} for error: {error_code}"
        )

    def degrade_feature(self, feature: str, reason: str):
        """Mark a feature as degraded."""
        self.degraded_features.add(feature)
        self.logger.warning(f"Feature degraded: {feature} - {reason}")

    def restore_feature(self, feature: str):
        """Restore a previously degraded feature."""
        if feature in self.degraded_features:
            self.degraded_features.remove(feature)
            self.logger.info(f"Feature restored: {feature}")

    def is_feature_degraded(self, feature: str) -> bool:
        """Check if a feature is currently degraded."""
        return feature in self.degraded_features

    async def execute_with_fallback(
        self, operation: str, primary_func: Callable, *args, **kwargs
    ) -> Any:
        """Execute an operation with fallback support."""
        try:
            # Try primary operation
            if asyncio.iscoroutinefunction(primary_func):
                return await primary_func(*args, **kwargs)
            else:
                return primary_func(*args, **kwargs)

        except Exception as e:
            self.logger.warning(f"Primary operation {operation} failed: {e}")

            # Try fallback if available
            if operation in self.fallback_handlers:
                try:
                    fallback_func = self.fallback_handlers[operation]
                    self.logger.info(f"Attempting fallback for operation: {operation}")

                    if asyncio.iscoroutinefunction(fallback_func):
                        result = await fallback_func(*args, **kwargs)
                    else:
                        result = fallback_func(*args, **kwargs)

                    self.logger.info(f"Fallback successful for operation: {operation}")
                    return result

                except Exception as fallback_error:
                    self.logger.error(
                        f"Fallback also failed for {operation}: {fallback_error}"
                    )
                    raise fallback_error

            # No fallback available, re-raise original error
            raise e

    def get_degradation_status(self) -> Dict[str, Any]:
        """Get current degradation status."""
        return {
            "degraded_features": list(self.degraded_features),
            "available_fallbacks": list(self.fallback_handlers.keys()),
            "recovery_strategies": {
                code: info["strategy"].value
                for code, info in self.recovery_strategies.items()
            },
        }


class RetryManager:
    """Manages retry logic with exponential backoff."""

    def __init__(self, config: RecoveryConfig, logger: Optional[logging.Logger] = None):
        self.config = config
        self.logger = logger or get_logger("retry")

    async def retry_async(self, func: Callable, *args, **kwargs) -> Any:
        """Retry an async function with exponential backoff."""
        last_exception = None
        delay = self.config.retry_delay

        for attempt in range(self.config.max_retries + 1):
            try:
                if attempt > 0:
                    self.logger.info(
                        f"Retry attempt {attempt}/{self.config.max_retries}"
                    )

                if asyncio.iscoroutinefunction(func):
                    return await asyncio.wait_for(
                        func(*args, **kwargs), timeout=self.config.timeout
                    )
                else:
                    return func(*args, **kwargs)

            except Exception as e:
                last_exception = e

                if attempt < self.config.max_retries:
                    self.logger.warning(
                        f"Attempt {attempt + 1} failed: {e}, retrying in {delay}s"
                    )
                    await asyncio.sleep(delay)
                    delay = min(
                        delay * self.config.backoff_multiplier, self.config.max_delay
                    )
                else:
                    self.logger.error(f"All retry attempts failed. Last error: {e}")

        raise last_exception

    def retry_sync(self, func: Callable, *args, **kwargs) -> Any:
        """Retry a synchronous function with exponential backoff."""
        last_exception = None
        delay = self.config.retry_delay

        for attempt in range(self.config.max_retries + 1):
            try:
                if attempt > 0:
                    self.logger.info(
                        f"Retry attempt {attempt}/{self.config.max_retries}"
                    )

                return func(*args, **kwargs)

            except Exception as e:
                last_exception = e

                if attempt < self.config.max_retries:
                    self.logger.warning(
                        f"Attempt {attempt + 1} failed: {e}, retrying in {delay}s"
                    )
                    time.sleep(delay)
                    delay = min(
                        delay * self.config.backoff_multiplier, self.config.max_delay
                    )
                else:
                    self.logger.error(f"All retry attempts failed. Last error: {e}")

        raise last_exception


class CircuitBreaker:
    """Circuit breaker pattern implementation for failing operations."""

    def __init__(
        self,
        failure_threshold: int = 5,
        recovery_timeout: float = 60.0,
        logger: Optional[logging.Logger] = None,
    ):
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.logger = logger or get_logger("circuit_breaker")

        self.failure_count = 0
        self.last_failure_time = None
        self.state = "closed"  # closed, open, half_open

    def __call__(self, func: Callable):
        """Decorator to wrap functions with circuit breaker."""

        async def async_wrapper(*args, **kwargs):
            if self.state == "open":
                if time.time() - self.last_failure_time > self.recovery_timeout:
                    self.state = "half_open"
                    self.logger.info("Circuit breaker transitioning to half-open state")
                else:
                    raise OperationError(
                        "Circuit breaker is open - operation blocked",
                        error_code=ErrorCode.OPERATION_FAILED.value,
                    )

            try:
                if asyncio.iscoroutinefunction(func):
                    result = await func(*args, **kwargs)
                else:
                    result = func(*args, **kwargs)

                # Success - reset failure count
                if self.state == "half_open":
                    self.state = "closed"
                    self.logger.info("Circuit breaker closed - operation successful")

                self.failure_count = 0
                return result

            except Exception as e:
                self.failure_count += 1
                self.last_failure_time = time.time()

                if self.failure_count >= self.failure_threshold:
                    self.state = "open"
                    self.logger.warning(
                        f"Circuit breaker opened after {self.failure_count} failures"
                    )

                raise e

        def sync_wrapper(*args, **kwargs):
            if self.state == "open":
                if time.time() - self.last_failure_time > self.recovery_timeout:
                    self.state = "half_open"
                    self.logger.info("Circuit breaker transitioning to half-open state")
                else:
                    raise OperationError(
                        "Circuit breaker is open - operation blocked",
                        error_code=ErrorCode.OPERATION_FAILED.value,
                    )

            try:
                result = func(*args, **kwargs)

                # Success - reset failure count
                if self.state == "half_open":
                    self.state = "closed"
                    self.logger.info("Circuit breaker closed - operation successful")

                self.failure_count = 0
                return result

            except Exception as e:
                self.failure_count += 1
                self.last_failure_time = time.time()

                if self.failure_count >= self.failure_threshold:
                    self.state = "open"
                    self.logger.warning(
                        f"Circuit breaker opened after {self.failure_count} failures"
                    )

                raise e

        if asyncio.iscoroutinefunction(func):
            return async_wrapper
        else:
            return sync_wrapper

    def get_state(self) -> Dict[str, Any]:
        """Get current circuit breaker state."""
        return {
            "state": self.state,
            "failure_count": self.failure_count,
            "failure_threshold": self.failure_threshold,
            "last_failure_time": self.last_failure_time,
            "recovery_timeout": self.recovery_timeout,
        }


class HealthChecker:
    """Monitors system health and triggers recovery actions."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or get_logger("health_checker")
        self.health_checks = {}
        self.health_status = {}

    def register_health_check(
        self,
        name: str,
        check_func: Callable,
        interval: float = 30.0,
        timeout: float = 10.0,
    ):
        """Register a health check function."""
        self.health_checks[name] = {
            "func": check_func,
            "interval": interval,
            "timeout": timeout,
            "last_check": 0,
            "consecutive_failures": 0,
        }
        self.health_status[name] = {"healthy": True, "last_error": None}
        self.logger.debug(f"Registered health check: {name}")

    async def run_health_checks(self):
        """Run all registered health checks."""
        current_time = time.time()

        for name, check_info in self.health_checks.items():
            if current_time - check_info["last_check"] >= check_info["interval"]:
                await self._run_single_check(name, check_info)
                check_info["last_check"] = current_time

    async def _run_single_check(self, name: str, check_info: Dict[str, Any]):
        """Run a single health check."""
        try:
            check_func = check_info["func"]

            if asyncio.iscoroutinefunction(check_func):
                result = await asyncio.wait_for(
                    check_func(), timeout=check_info["timeout"]
                )
            else:
                result = check_func()

            # Health check passed
            if check_info["consecutive_failures"] > 0:
                self.logger.info(f"Health check {name} recovered")

            check_info["consecutive_failures"] = 0
            self.health_status[name] = {"healthy": True, "last_error": None}

        except Exception as e:
            check_info["consecutive_failures"] += 1
            self.health_status[name] = {
                "healthy": False,
                "last_error": str(e),
                "consecutive_failures": check_info["consecutive_failures"],
            }

            self.logger.warning(
                f"Health check {name} failed (attempt {check_info['consecutive_failures']}): {e}"
            )

    def get_health_status(self) -> Dict[str, Any]:
        """Get current health status of all checks."""
        return {
            "overall_healthy": all(
                status["healthy"] for status in self.health_status.values()
            ),
            "checks": self.health_status.copy(),
        }

    def is_healthy(self, check_name: Optional[str] = None) -> bool:
        """Check if system or specific component is healthy."""
        if check_name:
            return self.health_status.get(check_name, {}).get("healthy", False)
        return all(status["healthy"] for status in self.health_status.values())


class RecoveryOrchestrator:
    """Orchestrates recovery actions across the system."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or get_logger("recovery_orchestrator")
        self.degradation = GracefulDegradation(logger)
        self.health_checker = HealthChecker(logger)
        self.recovery_actions = {}
        self.active_recoveries = set()

    def register_recovery_action(
        self,
        trigger: str,
        action_func: Callable,
        priority: int = 0,
        max_concurrent: int = 1,
    ):
        """Register a recovery action for a specific trigger."""
        self.recovery_actions[trigger] = {
            "func": action_func,
            "priority": priority,
            "max_concurrent": max_concurrent,
            "active_count": 0,
        }
        self.logger.debug(f"Registered recovery action for trigger: {trigger}")

    async def trigger_recovery(self, trigger: str, context: Dict[str, Any] = None):
        """Trigger recovery actions for a specific event."""
        if trigger not in self.recovery_actions:
            self.logger.warning(f"No recovery action registered for trigger: {trigger}")
            return

        action_info = self.recovery_actions[trigger]

        # Check if we can run this recovery action
        if action_info["active_count"] >= action_info["max_concurrent"]:
            self.logger.warning(f"Recovery action {trigger} already at max concurrency")
            return

        action_info["active_count"] += 1
        self.active_recoveries.add(trigger)

        try:
            self.logger.info(f"Starting recovery action: {trigger}")
            action_func = action_info["func"]

            if asyncio.iscoroutinefunction(action_func):
                await action_func(context or {})
            else:
                action_func(context or {})

            self.logger.info(f"Recovery action completed: {trigger}")

        except Exception as e:
            self.logger.error(f"Recovery action failed: {trigger} - {e}")

        finally:
            action_info["active_count"] -= 1
            self.active_recoveries.discard(trigger)

    async def monitor_and_recover(self, interval: float = 30.0):
        """Continuously monitor health and trigger recovery as needed."""
        while True:
            try:
                await self.health_checker.run_health_checks()

                # Check for unhealthy components and trigger recovery
                health_status = self.health_checker.get_health_status()
                for check_name, status in health_status["checks"].items():
                    if not status["healthy"]:
                        await self.trigger_recovery(
                            f"unhealthy_{check_name}",
                            {
                                "check_name": check_name,
                                "error": status.get("last_error"),
                                "consecutive_failures": status.get(
                                    "consecutive_failures", 0
                                ),
                            },
                        )

                await asyncio.sleep(interval)

            except Exception as e:
                self.logger.error(f"Error in monitor_and_recover: {e}")
                await asyncio.sleep(interval)

    def get_recovery_status(self) -> Dict[str, Any]:
        """Get current recovery system status."""
        return {
            "active_recoveries": list(self.active_recoveries),
            "registered_actions": list(self.recovery_actions.keys()),
            "degradation_status": self.degradation.get_degradation_status(),
            "health_status": self.health_checker.get_health_status(),
        }
