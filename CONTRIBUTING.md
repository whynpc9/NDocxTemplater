# Contributing

## Development setup

- Install .NET SDK 10.0+
- Restore/build:

```bash
dotnet build NDocxTemplater.sln --disable-build-servers -m:1
```

## Testing

```bash
dotnet test NDocxTemplater.sln --disable-build-servers -m:1
```

## Examples

Regenerate examples after template engine behavior changes:

```bash
dotnet run --project tools/ExampleGenerator/ExampleGenerator.csproj --disable-build-servers
```

## Pull requests

- Keep changes focused and include tests for behavior changes.
- Update `README.md` and `examples/` when adding template features.
- For every new feature, you must:
	1. Add corresponding test cases and verify they pass.
	2. Add a corresponding example under `examples/`.
	3. Update relevant documentation (at minimum `README.md`).
