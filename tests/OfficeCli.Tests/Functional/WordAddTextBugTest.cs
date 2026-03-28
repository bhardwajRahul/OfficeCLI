// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Regression test for the Text constructor bug: new Text(value) passes value as outerXml
/// instead of text content. Correct usage is new Text { Text = value }.
/// </summary>
public class WordAddTextBugTest : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public WordAddTextBugTest()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Add_Run_WithText_TextIsPersisted()
    {
        // Add paragraph then run with text
        _handler.Add("/", "p", null, new());
        _handler.Add("/body/p[1]", "r", null, new() { ["text"] = "Hello World" });

        var node = _handler.Get("/body/p[1]/r[1]");
        node.Text.Should().Be("Hello World");

        // Verify persistence after reopen
        Reopen();
        var node2 = _handler.Get("/body/p[1]/r[1]");
        node2.Text.Should().Be("Hello World");
    }

    [Fact]
    public void Add_Run_WithFormattedText_TextAndFormatPersisted()
    {
        _handler.Add("/", "p", null, new());
        _handler.Add("/body/p[1]", "r", null, new()
        {
            ["text"] = "Bold Red",
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        var node = _handler.Get("/body/p[1]/r[1]");
        node.Text.Should().Be("Bold Red");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void Add_Comment_TextIsPersisted()
    {
        _handler.Add("/", "p", null, new());
        _handler.Add("/body/p[1]", "r", null, new() { ["text"] = "content" });
        _handler.Add("/body/p[1]/r[1]", "comment", null, new()
        {
            ["text"] = "This is a comment",
            ["author"] = "Test"
        });

        var comments = _handler.Query("comment");
        comments.Should().NotBeEmpty();
        comments[0].Text.Should().Be("This is a comment");
    }

    [Fact]
    public void Add_Hyperlink_TextIsPersisted()
    {
        _handler.Add("/", "p", null, new());
        _handler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Click here"
        });

        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node.Text.Should().Contain("Click here");
    }

    [Fact]
    public void Add_TableCell_TextIsPersisted()
    {
        _handler.Add("/", "tbl", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell A1" });

        var node = _handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Contain("Cell A1");
    }

    [Fact]
    public void Add_Footnote_TextIsPersisted()
    {
        _handler.Add("/", "p", null, new());
        _handler.Add("/body/p[1]", "r", null, new() { ["text"] = "text" });
        _handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "My footnote" });

        var fn = _handler.Get("/footnote[1]");
        fn.Text.Should().Contain("My footnote");
    }
}
