#!/usr/bin/env bash
set -euo pipefail

dotnet run --project ../../tools/NDocxTemplater.Cli/NDocxTemplater.Cli.csproj -- \
  inspect-tags --template template.docx

dotnet run --project ../../tools/NDocxTemplater.Cli/NDocxTemplater.Cli.csproj -- \
  validate --template template.docx --data data.json

dotnet run --project ../../tools/NDocxTemplater.Cli/NDocxTemplater.Cli.csproj -- \
  render --template template.docx --data data.json --output output.docx
