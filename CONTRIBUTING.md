# Contributing to xlPort

Thank you for your interest in contributing to xlPort!

## How to Contribute

1. **Fork** the repository and create a branch from `master`.
2. **Make your changes.** Follow the existing code style (enforced by `google-java-format`).
3. **Add tests** for any new functionality. Test suites live in `src/test/resources/test-suites/` -- each test is a folder with a template, request JSON, and expected output.
4. **Run the tests:** `mvn clean verify`
5. **Submit a pull request** with a clear description of what you changed and why.

## Reporting Issues

Open a GitHub issue with:
- A description of the problem or feature request.
- Steps to reproduce (for bugs).
- The version of Java and xlPort you are using.

## Development Setup

```bash
# Build and test
mvn clean verify

# Run locally on port 8082
mvn jetty:run
```

## Code Style

This project uses [google-java-format](https://github.com/google/google-java-format). Please format your code before submitting.

## License

By contributing, you agree that your contributions will be licensed under the [Apache License 2.0](LICENSE).
