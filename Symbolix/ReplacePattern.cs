
// (c) 2023 Kazuki KOHZUKI

namespace Symbolix;

internal class ReplacePattern
{
    internal string Find { get; }

    internal string Replace { get; }

    internal ReplacePattern(string find, string replace)
    {
        this.Find = find;
        this.Replace = replace;
    } // ctor (string, string)
} // internal class ReplacePattern
