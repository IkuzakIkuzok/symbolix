
// (c) 2023 Kazuki KOHZUKI

using System;
using System.Collections.Generic;
using System.IO;
using ReplacePatterns = System.Collections.Generic.List<Symbolix.ReplacePattern>;

namespace Symbolix;

#nullable enable

internal sealed class Config
{
    private const string FileName = ".symbolixconfig";

    private static readonly string primaryConfigPath;

    internal bool CheckMinus { get; private set; } = true;
    internal ReplacePatterns SimpleReplacement { get; private set; } = new();
    internal List<string> MustBeMinus { get; private set; } = new();
    internal List<string> MustBeEnDash { get; private set; } = new();
    internal List<string> MustBeEmDash { get; private set; } = new();

    static Config()
    {
        var dir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        primaryConfigPath = Path.Combine(dir, FileName);
    } // cctor ()

    private Config() { }

    internal static Config Load(string directory)
    {
        var config = new Config();

        if (File.Exists(primaryConfigPath))
            config.LoadConfig(primaryConfigPath);

        var dir = directory;
        var files = new List<string>();
        while (!string.IsNullOrEmpty(dir))
        {
            var filename = Path.Combine(dir, FileName);
            if (File.Exists(filename)) files.Add(filename);
            dir = Path.GetDirectoryName(dir);
        }

        files.Reverse();
        foreach (var filename in files)
            config.LoadConfig(filename);

        return config;
    } // internal static Config Load (string)

    private void LoadConfig(string filename)
    {
        if (!File.Exists(filename)) return;

        var jsonText = File.ReadAllText(filename);
        var json = ConfigJson.LoadJson(jsonText);
        if (json == null) return;

        if (json.CheckMinus != null) this.CheckMinus = json.CheckMinus.Value;

        foreach (var replace in json.Replace ?? Array.Empty<ConfigJson.JsonReplaceElementObject>())
        {
            if (replace.Action == "add")
                this.SimpleReplacement.Add(new(replace.Find, replace.Replace));
            else if (replace.Action == "remove")
                this.SimpleReplacement.RemoveAll(x => x.Find == replace.Find);
        }

        void RegisterPattern(ConfigJson.JsonHyphenReplaceElementObject[]? patterns, List<string> list)
        {
            foreach (var pattern in patterns ?? Array.Empty<ConfigJson.JsonHyphenReplaceElementObject>())
            {
                var find = pattern.Find;
                if (string.IsNullOrEmpty(find)) continue;
                if (pattern.Action == "add")
                    list.Add(find!);
                else if (pattern.Action == "remove")
                    list.RemoveAll(x => x == find);
            }
        }

        RegisterPattern(json.Minus, this.MustBeMinus);
        RegisterPattern(json.EnDash, this.MustBeEnDash);
        RegisterPattern(json.EmDash, this.MustBeEmDash);
    } // private void LoadConfig (string)
} // internal sealed class Config
