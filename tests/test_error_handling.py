"""
Unit tests for error handling and recovery mechanisms.
"""

import pytest
import asyncio
import logging
from unittest.mock import Mock, patch, MagicMock
from typing import Dict, Any

from word_mcp_server.utils.errors import (
    WordMCPError, ConnectionError, DocumentError, OperationError, TimeoutError,
    ConfigurationError, ErrorCode, ErrorSuggestions, ErrorHandler
)
from word_mcp_server.utils.recovery import (
    GracefulDegradation, RetryManager, CircuitBreaker, HealthChecker,
    RecoveryOrchestrator, RecoveryConfig, RecoveryStrategy
)


class TestWordMCPError:
    """Test cases for WordMCPError and related exception classes."""
    
    def test_basic_error_creation(self):
        """Test basic error creation with message."""
        error = WordMCPError("Test error message")
        
        assert error.message == "Test error message"
        assert error.error_code == ErrorCode.UNKNOWN_ERROR.value
        assert error.details is None
        assert isinstance(error.suggestions, list)
        assert len(error.suggestions) > 0
    
    def test_error_with_all_parameters(self):
        """Test error creation with all parameters."""
        context = {"doc_id": "test-123", "operation": "save"}
        suggestions = ["Try again", "Check permissions"]
        
        error = WordMCPError(
            message="Custom error",
            error_code=ErrorCode.DOCUMENT_ACCESS_DENIED.value,
            details="File is locked by another process",
            suggestions=suggestions,
            context=context
        )
        
        assert error.message == "Custom error"
        assert error.error_code == ErrorCode.DOCUMENT_ACCESS_DENIED.value
        assert error.details == "File is locked by another process"
        assert error.suggestions == suggestions
        assert error.context == context
    
    def test_error_auto_suggestions(self):
        """Test automatic suggestion generation."""
        error = WordMCPError(
            "Connection failed",
            error_code=ErrorCode.WORD_CONNECTION_FAILED.value
        )
        
        suggestions = ErrorSuggestions.get_suggestions(ErrorCode.WORD_CONNECTION_FAILED)
        assert error.suggestions == suggestions
        assert "Ensure Microsoft Word is installed on your system" in error.suggestions
    
    def test_error_to_dict(self):
        """Test error serialization to dictionary."""
        context = {"doc_id": "test-123"}
        error = WordMCPError(
            "Test error",
            error_code=ErrorCode.OPERATION_FAILED.value,
            details="Operation details",
            context=context
        )
        
        error_dict = error.to_dict()
        
        assert error_dict["error"]["code"] == ErrorCode.OPERATION_FAILED.value
        assert error_dict["error"]["message"] == "Test error"
        assert error_dict["error"]["details"] == "Operation details"
        assert error_dict["error"]["context"] == context
        assert isinstance(error_dict["error"]["suggestions"], list)
    
    def test_error_logging(self):
        """Test error logging functionality."""
        mock_logger = Mock(spec=logging.Logger)
        
        error = WordMCPError(
            "Test error",
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            details="File not found"
        )
        
        error.log_error(mock_logger)
        
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args
        assert "Error DOCUMENT_NOT_FOUND: Test error" in call_args[0][0]
        assert call_args[1]["extra"]["error_code"] == ErrorCode.DOCUMENT_NOT_FOUND.value


class TestSpecificErrors:
    """Test cases for specific error types."""
    
    def test_connection_error(self):
        """Test ConnectionError creation."""
        error = ConnectionError("Failed to connect")
        
        assert error.error_code == ErrorCode.WORD_CONNECTION_FAILED.value
        assert error.message == "Failed to connect"
        assert "Ensure Microsoft Word is installed on your system" in error.suggestions
    
    def test_document_error(self):
        """Test DocumentError creation."""
        error = DocumentError(
            "Document not found",
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value
        )
        
        assert error.error_code == ErrorCode.DOCUMENT_NOT_FOUND.value
        assert error.message == "Document not found"
    
    def test_operation_error(self):
        """Test OperationError creation."""
        error = OperationError("Operation failed")
        
        assert error.error_code == ErrorCode.OPERATION_FAILED.value
        assert error.message == "Operation failed"
    
    def test_timeout_error(self):
        """Test TimeoutError creation."""
        error = TimeoutError("Operation timed out")
        
        assert error.error_code == ErrorCode.TIMEOUT_ERROR.value
        assert error.message == "Operation timed out"
    
    def test_configuration_error(self):
        """Test ConfigurationError creation."""
        error = ConfigurationError("Invalid config")
        
        assert error.error_code == ErrorCode.CONFIGURATION_ERROR.value
        assert error.message == "Invalid config"


class TestErrorSuggestions:
    """Test cases for error suggestion system."""
    
    def test_get_suggestions_for_known_error(self):
        """Test getting suggestions for known error codes."""
        suggestions = ErrorSuggestions.get_suggestions(ErrorCode.WORD_CONNECTION_FAILED)
        
        assert isinstance(suggestions, list)
        assert len(suggestions) > 0
        assert "Ensure Microsoft Word is installed on your system" in suggestions
    
    def test_get_suggestions_for_unknown_error(self):
        """Test getting suggestions for unknown error codes."""
        # Create a mock error code that doesn't exist in suggestions
        mock_error = Mock()
        mock_error.value = "UNKNOWN_TEST_ERROR"
        
        suggestions = ErrorSuggestions.get_suggestions(mock_error)
        
        assert isinstance(suggestions, list)
        assert len(suggestions) > 0
        assert "Check the application logs for more details" in suggestions
    
    def test_all_error_codes_have_suggestions(self):
        """Test that all defined error codes have suggestions."""
        for error_code in ErrorCode:
            suggestions = ErrorSuggestions.get_suggestions(error_code)
            assert isinstance(suggestions, list)
            assert len(suggestions) > 0


class TestErrorHandler:
    """Test cases for centralized error handler."""
    
    @pytest.fixture
    def mock_logger(self):
        """Create a mock logger."""
        return Mock(spec=logging.Logger)
    
    @pytest.fixture
    def error_handler(self, mock_logger):
        """Create an ErrorHandler instance."""
        return ErrorHandler(mock_logger)
    
    def test_handle_word_mcp_error(self, error_handler):
        """Test handling WordMCPError."""
        error = WordMCPError(
            "Test error",
            error_code=ErrorCode.OPERATION_FAILED.value
        )
        
        result = error_handler.handle_error(error)
        
        assert result["error"]["code"] == ErrorCode.OPERATION_FAILED.value
        assert result["error"]["message"] == "Test error"
        assert error_handler.error_counts[ErrorCode.OPERATION_FAILED.value] == 1
    
    def test_handle_generic_exception(self, error_handler):
        """Test handling generic Python exceptions."""
        error = ValueError("Generic error")
        context = {"operation": "test"}
        
        result = error_handler.handle_error(error, context)
        
        assert result["error"]["code"] == ErrorCode.UNKNOWN_ERROR.value
        assert result["error"]["message"] == "Generic error"
        assert result["error"]["context"] == context
    
    def test_error_frequency_tracking(self, error_handler):
        """Test error frequency tracking."""
        error_code = ErrorCode.DOCUMENT_NOT_FOUND.value
        
        # Generate multiple errors
        for i in range(3):
            error = DocumentError("Test", error_code=error_code)
            error_handler.handle_error(error)
        
        stats = error_handler.get_error_statistics()
        assert stats[error_code] == 3
    
    def test_frequent_error_warning(self, error_handler, mock_logger):
        """Test warning for frequent errors."""
        error_code = ErrorCode.OPERATION_FAILED.value
        
        # Generate enough errors to trigger warning
        for i in range(6):
            error = OperationError("Test")
            error_handler.handle_error(error)
        
        # Check that warning was logged
        warning_calls = [call for call in mock_logger.warning.call_args_list 
                        if "Frequent error detected" in str(call)]
        assert len(warning_calls) > 0
    
    def test_recovery_strategy_registration(self, error_handler):
        """Test recovery strategy registration."""
        def mock_recovery(error):
            return {"recovered": True}
        
        error_code = ErrorCode.WORD_CONNECTION_FAILED.value
        error_handler.register_recovery_strategy(error_code, mock_recovery)
        
        assert error_code in error_handler.recovery_strategies
        assert error_handler.recovery_strategies[error_code] == mock_recovery
    
    def test_recovery_strategy_execution(self, error_handler):
        """Test recovery strategy execution."""
        def mock_recovery(error):
            return {"recovered": True, "method": "restart"}
        
        error_code = ErrorCode.WORD_CONNECTION_FAILED.value
        error_handler.register_recovery_strategy(error_code, mock_recovery)
        
        error = ConnectionError("Test connection error")
        result = error_handler.handle_error(error)
        
        assert "recovery" in result
        assert result["recovery"]["recovered"] is True
        assert result["recovery"]["method"] == "restart"
    
    def test_recovery_strategy_failure(self, error_handler, mock_logger):
        """Test recovery strategy failure handling."""
        def failing_recovery(error):
            raise Exception("Recovery failed")
        
        error_code = ErrorCode.WORD_CONNECTION_FAILED.value
        error_handler.register_recovery_strategy(error_code, failing_recovery)
        
        error = ConnectionError("Test connection error")
        result = error_handler.handle_error(error)
        
        # Should not have recovery in result if it failed
        assert "recovery" not in result
        
        # Should log warning about recovery failure
        warning_calls = [call for call in mock_logger.warning.call_args_list 
                        if "Recovery strategy failed" in str(call)]
        assert len(warning_calls) > 0
    
    def test_reset_error_statistics(self, error_handler):
        """Test resetting error statistics."""
        error = OperationError("Test")
        error_handler.handle_error(error)
        
        assert len(error_handler.get_error_statistics()) > 0
        
        error_handler.reset_error_statistics()
        assert len(error_handler.get_error_statistics()) == 0


class TestGracefulDegradation:
    """Test cases for graceful degradation system."""
    
    @pytest.fixture
    def degradation(self):
        """Create a GracefulDegradation instance."""
        return GracefulDegradation()
    
    def test_feature_degradation(self, degradation):
        """Test feature degradation and restoration."""
        feature = "document_creation"
        reason = "Word connection lost"
        
        # Initially not degraded
        assert not degradation.is_feature_degraded(feature)
        
        # Degrade feature
        degradation.degrade_feature(feature, reason)
        assert degradation.is_feature_degraded(feature)
        
        # Restore feature
        degradation.restore_feature(feature)
        assert not degradation.is_feature_degraded(feature)
    
    def test_fallback_registration(self, degradation):
        """Test fallback handler registration."""
        def mock_fallback(*args, **kwargs):
            return "fallback_result"
        
        operation = "create_document"
        degradation.register_fallback(operation, mock_fallback)
        
        assert operation in degradation.fallback_handlers
        assert degradation.fallback_handlers[operation] == mock_fallback
    
    @pytest.mark.asyncio
    async def test_execute_with_fallback_success(self, degradation):
        """Test successful primary operation execution."""
        async def primary_func(value):
            return f"primary_{value}"
        
        result = await degradation.execute_with_fallback(
            "test_op", primary_func, "test"
        )
        
        assert result == "primary_test"
    
    @pytest.mark.asyncio
    async def test_execute_with_fallback_primary_fails(self, degradation):
        """Test fallback execution when primary fails."""
        async def primary_func(value):
            raise Exception("Primary failed")
        
        async def fallback_func(value):
            return f"fallback_{value}"
        
        degradation.register_fallback("test_op", fallback_func)
        
        result = await degradation.execute_with_fallback(
            "test_op", primary_func, "test"
        )
        
        assert result == "fallback_test"
    
    @pytest.mark.asyncio
    async def test_execute_with_fallback_both_fail(self, degradation):
        """Test when both primary and fallback fail."""
        async def primary_func(value):
            raise ValueError("Primary failed")
        
        async def fallback_func(value):
            raise RuntimeError("Fallback failed")
        
        degradation.register_fallback("test_op", fallback_func)
        
        with pytest.raises(RuntimeError, match="Fallback failed"):
            await degradation.execute_with_fallback(
                "test_op", primary_func, "test"
            )
    
    @pytest.mark.asyncio
    async def test_execute_with_fallback_no_fallback(self, degradation):
        """Test when primary fails and no fallback is registered."""
        async def primary_func(value):
            raise ValueError("Primary failed")
        
        with pytest.raises(ValueError, match="Primary failed"):
            await degradation.execute_with_fallback(
                "test_op", primary_func, "test"
            )
    
    def test_degradation_status(self, degradation):
        """Test getting degradation status."""
        # Add some test data
        degradation.degrade_feature("feature1", "reason1")
        degradation.register_fallback("op1", lambda: None)
        degradation.register_recovery_strategy(
            ErrorCode.WORD_CONNECTION_FAILED.value, 
            RecoveryStrategy.RETRY
        )
        
        status = degradation.get_degradation_status()
        
        assert "feature1" in status["degraded_features"]
        assert "op1" in status["available_fallbacks"]
        assert ErrorCode.WORD_CONNECTION_FAILED.value in status["recovery_strategies"]


class TestRetryManager:
    """Test cases for retry manager."""
    
    @pytest.fixture
    def retry_config(self):
        """Create a retry configuration."""
        return RecoveryConfig(
            max_retries=3,
            retry_delay=0.1,  # Short delay for tests
            backoff_multiplier=2.0,
            max_delay=1.0,
            timeout=5.0
        )
    
    @pytest.fixture
    def retry_manager(self, retry_config):
        """Create a RetryManager instance."""
        return RetryManager(retry_config)
    
    @pytest.mark.asyncio
    async def test_retry_async_success_first_attempt(self, retry_manager):
        """Test successful async operation on first attempt."""
        async def success_func():
            return "success"
        
        result = await retry_manager.retry_async(success_func)
        assert result == "success"
    
    @pytest.mark.asyncio
    async def test_retry_async_success_after_retries(self, retry_manager):
        """Test successful async operation after retries."""
        call_count = 0
        
        async def eventually_success_func():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise Exception(f"Attempt {call_count} failed")
            return "success"
        
        result = await retry_manager.retry_async(eventually_success_func)
        assert result == "success"
        assert call_count == 3
    
    @pytest.mark.asyncio
    async def test_retry_async_all_attempts_fail(self, retry_manager):
        """Test async operation that fails all attempts."""
        async def always_fail_func():
            raise ValueError("Always fails")
        
        with pytest.raises(ValueError, match="Always fails"):
            await retry_manager.retry_async(always_fail_func)
    
    def test_retry_sync_success_first_attempt(self, retry_manager):
        """Test successful sync operation on first attempt."""
        def success_func():
            return "success"
        
        result = retry_manager.retry_sync(success_func)
        assert result == "success"
    
    def test_retry_sync_success_after_retries(self, retry_manager):
        """Test successful sync operation after retries."""
        call_count = 0
        
        def eventually_success_func():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise Exception(f"Attempt {call_count} failed")
            return "success"
        
        result = retry_manager.retry_sync(eventually_success_func)
        assert result == "success"
        assert call_count == 3
    
    def test_retry_sync_all_attempts_fail(self, retry_manager):
        """Test sync operation that fails all attempts."""
        def always_fail_func():
            raise ValueError("Always fails")
        
        with pytest.raises(ValueError, match="Always fails"):
            retry_manager.retry_sync(always_fail_func)


class TestCircuitBreaker:
    """Test cases for circuit breaker."""
    
    @pytest.fixture
    def circuit_breaker(self):
        """Create a CircuitBreaker instance."""
        return CircuitBreaker(
            failure_threshold=3,
            recovery_timeout=1.0  # Short timeout for tests
        )
    
    def test_circuit_breaker_closed_state(self, circuit_breaker):
        """Test circuit breaker in closed state."""
        @circuit_breaker
        def success_func():
            return "success"
        
        result = success_func()
        assert result == "success"
        assert circuit_breaker.state == "closed"
    
    def test_circuit_breaker_opens_after_failures(self, circuit_breaker):
        """Test circuit breaker opens after threshold failures."""
        @circuit_breaker
        def failing_func():
            raise Exception("Function failed")
        
        # Cause enough failures to open circuit
        for i in range(3):
            with pytest.raises(Exception):
                failing_func()
        
        assert circuit_breaker.state == "open"
        
        # Next call should be blocked
        with pytest.raises(OperationError, match="Circuit breaker is open"):
            failing_func()
    
    @pytest.mark.asyncio
    async def test_circuit_breaker_async_function(self, circuit_breaker):
        """Test circuit breaker with async functions."""
        @circuit_breaker
        async def async_func():
            return "async_success"
        
        result = await async_func()
        assert result == "async_success"
    
    def test_circuit_breaker_state_info(self, circuit_breaker):
        """Test getting circuit breaker state information."""
        state = circuit_breaker.get_state()
        
        assert state["state"] == "closed"
        assert state["failure_count"] == 0
        assert state["failure_threshold"] == 3
        assert state["recovery_timeout"] == 1.0


class TestHealthChecker:
    """Test cases for health checker."""
    
    @pytest.fixture
    def health_checker(self):
        """Create a HealthChecker instance."""
        return HealthChecker()
    
    def test_register_health_check(self, health_checker):
        """Test registering a health check."""
        def mock_check():
            return True
        
        health_checker.register_health_check("test_check", mock_check)
        
        assert "test_check" in health_checker.health_checks
        assert "test_check" in health_checker.health_status
        assert health_checker.health_status["test_check"]["healthy"] is True
    
    @pytest.mark.asyncio
    async def test_run_health_checks_success(self, health_checker):
        """Test running successful health checks."""
        def successful_check():
            return True
        
        health_checker.register_health_check("success_check", successful_check, interval=0.1)
        
        await health_checker.run_health_checks()
        
        assert health_checker.health_status["success_check"]["healthy"] is True
    
    @pytest.mark.asyncio
    async def test_run_health_checks_failure(self, health_checker):
        """Test running failing health checks."""
        def failing_check():
            raise Exception("Health check failed")
        
        health_checker.register_health_check("fail_check", failing_check, interval=0.1)
        
        await health_checker.run_health_checks()
        
        assert health_checker.health_status["fail_check"]["healthy"] is False
        assert "Health check failed" in health_checker.health_status["fail_check"]["last_error"]
    
    def test_get_health_status(self, health_checker):
        """Test getting overall health status."""
        def healthy_check():
            return True
        
        def unhealthy_check():
            raise Exception("Unhealthy")
        
        health_checker.register_health_check("healthy", healthy_check)
        health_checker.register_health_check("unhealthy", unhealthy_check)
        
        # Manually set status for testing
        health_checker.health_status["healthy"] = {"healthy": True, "last_error": None}
        health_checker.health_status["unhealthy"] = {"healthy": False, "last_error": "Unhealthy"}
        
        status = health_checker.get_health_status()
        
        assert status["overall_healthy"] is False
        assert status["checks"]["healthy"]["healthy"] is True
        assert status["checks"]["unhealthy"]["healthy"] is False
    
    def test_is_healthy(self, health_checker):
        """Test checking if system is healthy."""
        health_checker.health_status["test1"] = {"healthy": True, "last_error": None}
        health_checker.health_status["test2"] = {"healthy": False, "last_error": "Error"}
        
        assert not health_checker.is_healthy()  # Overall not healthy
        assert health_checker.is_healthy("test1")  # Specific check healthy
        assert not health_checker.is_healthy("test2")  # Specific check unhealthy


class TestRecoveryOrchestrator:
    """Test cases for recovery orchestrator."""
    
    @pytest.fixture
    def orchestrator(self):
        """Create a RecoveryOrchestrator instance."""
        return RecoveryOrchestrator()
    
    def test_register_recovery_action(self, orchestrator):
        """Test registering recovery actions."""
        def mock_recovery(context):
            return {"recovered": True}
        
        orchestrator.register_recovery_action("test_trigger", mock_recovery)
        
        assert "test_trigger" in orchestrator.recovery_actions
        assert orchestrator.recovery_actions["test_trigger"]["func"] == mock_recovery
    
    @pytest.mark.asyncio
    async def test_trigger_recovery_sync(self, orchestrator):
        """Test triggering synchronous recovery action."""
        recovery_called = False
        
        def mock_recovery(context):
            nonlocal recovery_called
            recovery_called = True
            assert context["test_key"] == "test_value"
        
        orchestrator.register_recovery_action("test_trigger", mock_recovery)
        
        await orchestrator.trigger_recovery("test_trigger", {"test_key": "test_value"})
        
        assert recovery_called
    
    @pytest.mark.asyncio
    async def test_trigger_recovery_async(self, orchestrator):
        """Test triggering asynchronous recovery action."""
        recovery_called = False
        
        async def mock_recovery(context):
            nonlocal recovery_called
            recovery_called = True
            assert context["test_key"] == "test_value"
        
        orchestrator.register_recovery_action("test_trigger", mock_recovery)
        
        await orchestrator.trigger_recovery("test_trigger", {"test_key": "test_value"})
        
        assert recovery_called
    
    @pytest.mark.asyncio
    async def test_trigger_recovery_unknown_trigger(self, orchestrator):
        """Test triggering recovery for unknown trigger."""
        # Should not raise exception, just log warning
        await orchestrator.trigger_recovery("unknown_trigger")
    
    @pytest.mark.asyncio
    async def test_trigger_recovery_concurrency_limit(self, orchestrator):
        """Test recovery action concurrency limits."""
        call_count = 0
        
        async def slow_recovery(context):
            nonlocal call_count
            call_count += 1
            await asyncio.sleep(0.1)  # Simulate slow recovery
        
        orchestrator.register_recovery_action("slow_trigger", slow_recovery, max_concurrent=1)
        
        # Start two recovery actions simultaneously
        task1 = asyncio.create_task(orchestrator.trigger_recovery("slow_trigger"))
        task2 = asyncio.create_task(orchestrator.trigger_recovery("slow_trigger"))
        
        await asyncio.gather(task1, task2)
        
        # Only one should have executed due to concurrency limit
        assert call_count == 1
    
    def test_get_recovery_status(self, orchestrator):
        """Test getting recovery system status."""
        def mock_recovery(context):
            pass
        
        orchestrator.register_recovery_action("test_action", mock_recovery)
        
        status = orchestrator.get_recovery_status()
        
        assert "active_recoveries" in status
        assert "registered_actions" in status
        assert "degradation_status" in status
        assert "health_status" in status
        assert "test_action" in status["registered_actions"]