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
