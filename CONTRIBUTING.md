# Contributing to excelstream

Thank you for your interest in contributing to excelstream! This document provides guidelines for contributing to the project.

## Code of Conduct

Please be respectful and considerate in all interactions. We aim to foster an inclusive and welcoming community.

## How to Contribute

### Reporting Bugs

1. Check if the bug has already been reported in [Issues](https://github.com/KSD-CO/excelstream/issues)
2. If not, create a new issue with:
   - Clear title and description
   - Steps to reproduce
   - Expected vs actual behavior
   - Version information (Rust version, OS, excelstream version)
   - Sample code if applicable

### Suggesting Features

1. Check existing issues for similar feature requests
2. Create a new issue describing:
   - Use case and motivation
   - Proposed API design
   - Alternative approaches considered
   - Breaking changes (if any)

### Pull Requests

1. **Fork and Clone**
   ```bash
   git clone https://github.com/YOUR_USERNAME/excelstream.git
   cd excelstream
   ```

2. **Create a Branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/bug-description
   ```

3. **Make Changes**
   - Follow the coding style (run `cargo fmt`)
   - Add tests for new functionality
   - Update documentation
   - Keep commits focused and atomic

4. **Test Your Changes**
   ```bash
   # Format code
   cargo fmt

   # Run clippy
   cargo clippy --all-targets --all-features -- -D warnings

   # Run all tests
   cargo test --all-features

   # Run examples
   cargo run --example basic_write
   ```

5. **Commit**
   ```bash
   git add .
   git commit -m "feat: add support for cell merging"
   # or
   git commit -m "fix: handle empty strings correctly"
   ```

   Commit message format:
   - `feat:` - New feature
   - `fix:` - Bug fix
   - `docs:` - Documentation changes
   - `test:` - Test additions or changes
   - `refactor:` - Code refactoring
   - `perf:` - Performance improvements
   - `chore:` - Maintenance tasks

6. **Push and Create PR**
   ```bash
   git push origin feature/your-feature-name
   ```
   Then create a Pull Request on GitHub

## Development Guidelines

### Code Style

- Follow Rust standard formatting (`cargo fmt`)
- Use meaningful variable and function names
- Add comments for complex logic
- Keep functions focused and small
- Avoid `unwrap()` in library code (use proper error handling)

### Testing

- Add unit tests for new functions
- Add integration tests for new features
- Test edge cases (empty data, special characters, large datasets)
- Ensure all tests pass before submitting PR

### Documentation

- Add doc comments for public APIs
- Include examples in doc comments
- Update README.md if adding major features
- Add entries to CHANGELOG.md for user-facing changes

### Performance

- Benchmark performance-critical changes
- Maintain or improve current performance metrics
- Consider memory usage for streaming operations

## Project Structure

```
excelstream/
├── src/
│   ├── lib.rs              # Library entry point
│   ├── reader.rs           # Excel reading
│   ├── writer.rs           # Excel writing (wrapper)
│   ├── types.rs            # Type definitions
│   ├── error.rs            # Error types
│   └── fast_writer/        # High-performance writer
│       ├── mod.rs
│       ├── workbook.rs
│       ├── worksheet.rs
│       └── ...
├── tests/                  # Integration tests
├── examples/               # Usage examples
├── benches/                # Performance benchmarks
└── docs/                   # Additional documentation
```

## Areas for Contribution

### High Priority
- Cell formatting and styling
- Formula support improvements
- Cell merging
- Data validation
- More comprehensive tests

### Medium Priority
- Conditional formatting
- Charts support
- Images support
- Performance optimizations
- Better error messages

### Documentation
- More examples
- Tutorial guides
- API documentation improvements
- Migration guides

## Release Process

1. Update version in `Cargo.toml`
2. Update `CHANGELOG.md`
3. Run full test suite
4. Create git tag: `git tag -a v0.x.x -m "Release v0.x.x"`
5. Push tag: `git push origin v0.x.x`
6. GitHub Actions will automatically publish to crates.io

## Getting Help

- Open an issue for questions
- Check existing documentation in `/docs`
- Read examples in `/examples`
- Review tests in `/tests` for usage patterns

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to excelstream! 🦀
