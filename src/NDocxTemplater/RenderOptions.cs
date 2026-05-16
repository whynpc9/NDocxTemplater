using System;
using System.Globalization;

namespace NDocxTemplater;

public enum MissingValueBehavior
{
    Empty,
    KeepTag,
    Throw
}

public sealed class RenderOptions
{
    private CultureInfo _culture = CultureInfo.InvariantCulture;

    public MissingValueBehavior MissingValueBehavior { get; set; } = MissingValueBehavior.Empty;

    public CultureInfo Culture
    {
        get => _culture;
        set => _culture = value ?? throw new ArgumentNullException(nameof(value));
    }

    public string? BaseDirectory { get; set; }

    public Action<RenderWarning>? WarningHandler { get; set; }

    internal static RenderOptions Default { get; } = new RenderOptions();
}

public sealed class RenderWarning
{
    public RenderWarning(string code, string message, string? expression)
    {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Expression = expression;
    }

    public string Code { get; }

    public string Message { get; }

    public string? Expression { get; }
}
