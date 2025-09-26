# COM Resource Management Testing Suite

This directory contains comprehensive tests for the COM (Component Object Model) resource management improvements implemented for GitHub issue #66. These tests validate that memory leaks are prevented, timeouts are handled properly, and COM objects are cleaned up correctly.

## 🎯 Test Objectives

The test suite validates the following COM resource management improvements:

1. **COMResourceManager Context Manager**: Proper initialization, object tracking, and cleanup
2. **Enhanced Finally Blocks**: COM resource cleanup with garbage collection in Excel commands
3. **Improved HWP Export**: Better COM resource management for HWP operations
4. **Enhanced Timeout Handling**: Proper timeout management for pivot chart operations (Issue #42)

## 📁 Test File Structure

### Core Test Files

| Test File | Description | Focus Area |
|-----------|-------------|------------|
| `test_com_resource_manager.py` | Unit tests for COMResourceManager class | Context manager behavior, object tracking, API cleanup |
| `test_utils_timeout.py` | Tests for timeout utility functions | Timeout handling, thread management, COM cleanup |
| `test_excel_com_integration.py` | Integration tests for Excel command COM cleanup | Excel command integration, memory management |
| `test_com_performance_memory.py` | Performance and memory leak detection tests | Memory usage, leak prevention, scalability |
| `test_com_edge_cases.py` | Edge cases and error scenario tests | Exception handling, platform differences, error recovery |
| `run_com_tests.py` | Test runner and reporting utility | Automated test execution, coverage analysis |

### Supporting Files

- `conftest.py` - Pytest configuration and common fixtures
- `COM_TESTING_README.md` - This documentation file

## 🧪 Test Categories

### 1. Unit Tests (`test_com_resource_manager.py`)

**Coverage**: COMResourceManager class methods and behavior

**Key Test Scenarios**:
- ✅ Context manager enter/exit behavior
- ✅ COM object addition and tracking
- ✅ API reference management
- ✅ Garbage collection enforcement
- ✅ Windows-specific COM library cleanup
- ✅ Error handling during cleanup
- ✅ Verbose mode logging
- ✅ Object chaining support

**Sample Test**:
```python
def test_context_manager_exit_success(self):
    """컨텍스트 매니저 exit 성공 테스트"""
    manager = COMResourceManager()
    mock_obj = Mock()
    mock_obj.close = Mock()
    manager.add(mock_obj)

    with manager:
        pass

    mock_obj.close.assert_called_once()
    assert len(manager.com_objects) == 0
```

### 2. Timeout Utilities (`test_utils_timeout.py`)

**Coverage**: utils_timeout.py functions for timeout handling

**Key Test Scenarios**:
- ✅ Successful function execution with timeout
- ✅ Timeout occurrence and handling
- ✅ Function exceptions within timeout
- ✅ Pivot layout connection with timeout
- ✅ Pivot operation cleanup with timeout
- ✅ COM cleanup after timeout failures
- ✅ Thread management and daemon behavior

**Sample Test**:
```python
def test_timeout_occurrence(self):
    """타임아웃 발생 테스트"""
    def slow_func():
        time.sleep(2)
        return "should not reach here"

    success, result, error = execute_with_timeout(slow_func, timeout=1)

    assert success is False
    assert result is None
    assert "1초 내에 완료되지 않아 타임아웃" in error
```

### 3. Integration Tests (`test_excel_com_integration.py`)

**Coverage**: Excel command COM cleanup integration

**Key Test Scenarios**:
- ✅ Excel command COM cleanup in finally blocks
- ✅ Memory usage stability across multiple operations
- ✅ COMResourceManager integration with real commands
- ✅ Exception handling with COM cleanup
- ✅ Nested COM operations
- ✅ Timeout handling in chart operations

**Sample Test**:
```python
@patch('gc.collect')
def test_range_read_com_cleanup(self, mock_gc, mock_xlwings_with_com):
    """range-read 명령의 COM 정리 테스트"""
    runner = CliRunner()
    result = runner.invoke(excel_app, [
        "range-read", "--file-path", str(temp_path),
        "--range", "A1:B2", "--format", "json"
    ])

    assert result.exit_code == 0
    assert mock_gc.called  # finally 블록에서 호출됨
```

### 4. Performance & Memory Tests (`test_com_performance_memory.py`)

**Coverage**: Memory leak detection and performance validation

**Key Test Scenarios**:
- ✅ Memory leak prevention with repeated operations
- ✅ Large-scale COM object handling (1000+ objects)
- ✅ Concurrent COM operations safety
- ✅ Memory peak monitoring
- ✅ Garbage collection effectiveness
- ✅ Scalability with deep nesting
- ✅ High-frequency operation stability

**Sample Test**:
```python
def test_memory_leak_prevention_basic(self):
    """기본 메모리 누수 방지 테스트"""
    tracker = MemoryTracker()

    for iteration in range(10):
        with COMResourceManager() as com_manager:
            for i in range(5):
                large_obj = LargeCOMObject(0.5)  # 0.5MB 객체
                com_manager.add(large_obj)

    # 메모리 증가량이 5MB 미만이어야 함
    assert tracker.memory_increase < 5.0
```

### 5. Edge Cases (`test_com_edge_cases.py`)

**Coverage**: Error conditions and platform-specific behavior

**Key Test Scenarios**:
- ✅ COM objects with broken cleanup methods
- ✅ Missing API references and Release methods
- ✅ Platform-specific COM handling (Windows/macOS/Linux)
- ✅ Unicode and very long object names
- ✅ Concurrent access race conditions
- ✅ Resource exhaustion scenarios
- ✅ System limitation handling

**Sample Test**:
```python
def test_cleanup_with_broken_close_method(self):
    """close 메서드가 실패하는 객체 정리 테스트"""
    with COMResourceManager(verbose=True) as com_manager:
        broken_obj = BrokenCOMObject("close")
        com_manager.add(broken_obj)

    # 에러 발생해도 컨텍스트는 정상 종료
```

## 🚀 Running the Tests

### Prerequisites

```bash
# Install required testing dependencies
pip install pytest pytest-mock pytest-cov psutil

# For performance tests, also install:
pip install memory-profiler
```

### Quick Start

```bash
# Run all COM tests
python tests/run_com_tests.py all

# Run specific test categories
python tests/run_com_tests.py unit          # Unit tests only
python tests/run_com_tests.py timeout       # Timeout tests only
python tests/run_com_tests.py integration   # Integration tests only
python tests/run_com_tests.py performance   # Performance tests only
python tests/run_com_tests.py edge          # Edge case tests only

# Run with coverage analysis
python tests/run_com_tests.py coverage
```

### Individual Test Files

```bash
# Run individual test files with pytest
pytest tests/test_com_resource_manager.py -v
pytest tests/test_utils_timeout.py -v --tb=short
pytest tests/test_excel_com_integration.py -v -s
pytest tests/test_com_performance_memory.py -v -x
pytest tests/test_com_edge_cases.py -v

# Run with specific markers or patterns
pytest tests/ -k "test_timeout" -v
pytest tests/ -k "test_memory" -v
pytest tests/ -m "not slow" -v  # Skip slow tests (if marked)
```

### Coverage Analysis

```bash
# Generate coverage report
pytest tests/test_com_*.py --cov=pyhub_office_automation.excel --cov-report=html --cov-report=term-missing

# View coverage report
open htmlcov/index.html  # macOS
start htmlcov/index.html # Windows
```

## 📊 Expected Test Results

### Success Criteria

- **Unit Tests**: 100% pass rate (35+ test cases)
- **Timeout Tests**: 100% pass rate (25+ test cases)
- **Integration Tests**: 100% pass rate (20+ test cases)
- **Performance Tests**: 95%+ pass rate (memory-dependent)
- **Edge Cases**: 90%+ pass rate (system-dependent)
- **Overall Code Coverage**: 80%+ for COM management code

### Performance Benchmarks

| Test Category | Expected Time | Memory Limit |
|---------------|---------------|--------------|
| Unit Tests | < 30 seconds | < 50MB increase |
| Timeout Tests | < 60 seconds | < 20MB increase |
| Integration Tests | < 120 seconds | < 100MB increase |
| Performance Tests | < 300 seconds | < 200MB increase |
| Edge Cases | < 180 seconds | < 100MB increase |

### Platform-Specific Behavior

| Platform | COM Library Cleanup | Special Considerations |
|----------|-------------------|----------------------|
| Windows | ✅ pythoncom.CoUninitialize() | Full COM support |
| macOS | ❌ No COM cleanup needed | xlwings AppleScript mode |
| Linux | ❌ No COM cleanup needed | Excel unavailable |

## 🔧 Test Configuration

### Pytest Configuration

The tests use the following pytest configuration (in `conftest.py`):

```python
@pytest.fixture
def mock_xlwings_with_com():
    """COM 리소스 관리가 포함된 xlwings 모킹"""
    # Comprehensive xlwings mocking with COM objects

@pytest.fixture
def temp_excel_file():
    """임시 Excel 파일 생성"""
    # Temporary file management for tests
```

### Environment Variables

```bash
# Optional environment variables for test configuration
export PYTEST_TIMEOUT=300           # Test timeout in seconds
export COM_TEST_VERBOSE=1           # Enable verbose COM logging
export MEMORY_TEST_SIZE=100         # Adjust memory test object count
```

## 🐛 Debugging Test Failures

### Common Issues and Solutions

1. **Timeout Test Failures**
   ```bash
   # Increase timeout values for slow systems
   pytest tests/test_utils_timeout.py --timeout=600
   ```

2. **Memory Test Failures**
   ```bash
   # Run with reduced memory pressure
   pytest tests/test_com_performance_memory.py -s --tb=long
   ```

3. **Platform-Specific Failures**
   ```bash
   # Skip Windows-only tests on other platforms
   pytest tests/ -m "not windows_only"
   ```

4. **Integration Test Issues**
   ```bash
   # Run with mock debugging
   pytest tests/test_excel_com_integration.py -s -vv
   ```

### Debug Output

Enable verbose output for detailed debugging:

```python
# In test files, enable verbose mode
manager = COMResourceManager(verbose=True)

# Or via command line
pytest tests/ -s -vv --log-cli-level=DEBUG
```

## 📈 Test Metrics and Reporting

### Automated Reporting

The test runner generates comprehensive reports:

- **Execution Time**: Per-file and total execution time
- **Success Rate**: Percentage of passing tests
- **Memory Usage**: Peak memory increase during tests
- **Coverage Analysis**: Code coverage metrics
- **Platform Detection**: Automatic platform-specific test selection

### CI/CD Integration

For continuous integration, use:

```bash
# CI-friendly test execution
python tests/run_com_tests.py all --tb=short --no-header -q

# Generate JUnit XML for CI systems
pytest tests/test_com_*.py --junitxml=test-results.xml

# Memory profiling for CI
python -m memory_profiler tests/test_com_performance_memory.py
```

## 🎯 Test Development Guidelines

### Adding New Tests

1. **Follow Naming Convention**:
   - Test files: `test_com_*.py`
   - Test classes: `TestCOM*` or `Test*COM*`
   - Test methods: `test_*_com_*` for COM-specific tests

2. **Use Appropriate Fixtures**:
   ```python
   def test_new_com_feature(self, mock_xlwings_with_com):
       # Use provided fixtures for consistency
   ```

3. **Include Performance Considerations**:
   ```python
   def test_new_feature_performance(self):
       start_time = time.time()
       # ... test code ...
       assert (time.time() - start_time) < 1.0  # 1초 제한
   ```

4. **Add Platform Checks**:
   ```python
   @pytest.mark.skipif(platform.system() != "Windows", reason="COM은 Windows 전용")
   def test_windows_com_feature(self):
       # Windows-specific test
   ```

### Test Quality Standards

- **Isolation**: Each test should be independent
- **Deterministic**: Tests should produce consistent results
- **Fast**: Unit tests < 100ms, integration tests < 1s
- **Clear**: Descriptive test names and clear assertions
- **Comprehensive**: Cover both success and failure scenarios

## 📚 Related Documentation

- [GitHub Issue #66](https://github.com/pyhub-kr/pyhub-office-automation/issues/66) - COM Resource Management
- [GitHub Issue #42](https://github.com/pyhub-kr/pyhub-office-automation/issues/42) - Pivot Chart Timeout
- `specs/xlwings.md` - xlwings Integration Patterns
- `pyhub_office_automation/excel/utils.py` - COMResourceManager Implementation
- `pyhub_office_automation/excel/utils_timeout.py` - Timeout Utilities

## 🤝 Contributing

When contributing to the COM test suite:

1. **Run Full Test Suite**: Ensure all tests pass before submitting
2. **Update Documentation**: Update this README if adding new test categories
3. **Follow Patterns**: Use existing test patterns and fixtures
4. **Add Coverage**: Ensure new features are adequately tested
5. **Performance Impact**: Consider performance implications of new tests

---

**Last Updated**: 2025-09-25
**Test Suite Version**: 1.0
**Total Test Files**: 5
**Estimated Test Count**: 150+
**Expected Coverage**: 80%+