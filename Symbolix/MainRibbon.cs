
// (c) 2023 Kazuki KOHZUKI

using Microsoft.Office.Tools.Ribbon;
using Symbolix.Word;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Symbolix;

#nullable enable

[DesignerCategory("Code")]
public sealed class MainRibbon : RibbonBase
{
    private const string HYPHEN = "-";
    private const string MINUS = "−";
    private const string EN_DASH = "–";
    private const string EM_DASH = "—";

    private readonly IContainer? components = null;

    private RibbonTab tab;

    private RibbonGroup rg_run;
    private RibbonButton rb_run_thisDoc, rb_run_selection, rb_run_allDoc;

    public MainRibbon() : base(Globals.Factory.GetRibbonFactory())
    {
        InitializeComponent();
    } // ctor () : base (RibbonFactory)

    [MemberNotNull(
        nameof(this.tab), nameof(this.rg_run),
        nameof(this.rb_run_thisDoc), nameof(this.rb_run_selection), nameof(this.rb_run_allDoc)
    )]
    private void InitializeComponent()
    {
        this.tab = this.Factory.CreateRibbonTab();
        this.tab.Label = "Symbolix";
        this.tab.SuspendLayout();
        SuspendLayout();

        this.rg_run = this.Factory.CreateRibbonGroup();
        this.rg_run.Label = "Run";
        this.tab.Groups.Add(this.rg_run);

        this.rb_run_thisDoc = this.Factory.CreateRibbonButton();
        this.rb_run_thisDoc.Label = "This document";
        this.rb_run_thisDoc.Click += RunCheckThisDock;
        this.rg_run.Items.Add(this.rb_run_thisDoc);

        this.rb_run_selection = this.Factory.CreateRibbonButton();
        this.rb_run_selection.Label = "Selection";
        this.rb_run_selection.Click += RunCheckSelection;
        this.rg_run.Items.Add(this.rb_run_selection);

        this.rb_run_allDoc = this.Factory.CreateRibbonButton();
        this.rb_run_allDoc.Label = "All documents";
        this.rb_run_allDoc.Click += RunCheckAllDocs;
        this.rg_run.Items.Add(this.rb_run_allDoc);

        this.RibbonType = "Microsoft.Word.Document";
        this.Tabs.Add(this.tab);
        this.tab.ResumeLayout(false);
        this.tab.PerformLayout();
        ResumeLayout(false);
    } // private void InitializeComponent ()

    override protected void Dispose(bool disposing)
    {
        if (disposing) this.components?.Dispose();
        base.Dispose(disposing);
    } // override protected void Dispose (bool)

    private static void RunCheckThisDock(object? sender, RibbonControlEventArgs e)
        => RunCheckThisDock();

    private static void RunCheckThisDock()
    {
        var document = Document.ActiveDocument;
        if (document == null) return;

        RunCheck(document);
    } // private static void RunCheckThisDock ()

    private static void RunCheckSelection(object? sender, RibbonControlEventArgs e)
        => RunCheckThisPage();

    private static void RunCheckThisPage()
    {
        var selection = Range.Selection;
        RunCheck(selection);
    } // private static void RunCheckSelection ()

    private static void RunCheckAllDocs(object? sender, RibbonControlEventArgs e)
        => RunCheckAllDocs();

    private static void RunCheckAllDocs()
    {
        foreach (var document in Globals.ThisAddIn.Application.Documents.Cast<Document>())
            RunCheck(document);
    } // private static void RunCheckAllDocs ()

    private static void RunCheck(Document document)
        => RunCheck(document.Content);

    private static void RunCheck(Range range)
    {
        var document = range.Document;
        var trackRevisionsState = document.TrackRevisions;
        document.TrackRevisions = true;

        try
        {
            var config = Config.Load(document.Path);

            if (config.CheckMinus)
            {
                // Replace hyphens before numbers with minus signs.
                // Use a loop because wildcards (regular expressions) will cause the replacement result to be incorrect.
                for (var i = 0; i < 10; i++)
                {
                    var findObj = range.Find;
                    findObj.ReplaceAll($"{HYPHEN}{i}", $"{MINUS}{i}");
                }
            }

            foreach (var replacement in config.SimpleReplacement)
            {
                var find = replacement.Find;
                var replace = replacement.Replace;
                var findObj = range.Find;
                findObj.ReplaceAll(find, replace);
            }

            void ReplacePattern(List<string> patterns, string c)
            {
                foreach (var find in patterns)
                {
                    var replace = find.Replace(HYPHEN, c);
                    var range = document.Content;
                    var findObj = range.Find;
                    findObj.ReplaceAll(find, replace);
                }
            }

            ReplacePattern(config.MustBeMinus, MINUS);
            ReplacePattern(config.MustBeEnDash, EN_DASH);
            ReplacePattern(config.MustBeEmDash, EM_DASH);
        }
        finally
        {
            document.TrackRevisions = trackRevisionsState;
        }
    } // private static void RunCheck (Range)
} // public sealed class MainRibbon : RibbonBase
