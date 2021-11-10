﻿using System;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.Commanding;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Editor.Commanding;
using Microsoft.VisualStudio.Text.Editor.Commanding.Commands;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using ZenCoding;

namespace ZenCodingVS
{
    [Export(typeof(ICommandHandler))]
    [Name(nameof(Expander))]
    [ContentType("HTML")]
    [ContentType("HTMLX")]
    [ContentType("Razor")]
    [ContentType("LegacyRazorCSharp")]
    [TextViewRole(PredefinedTextViewRoles.PrimaryDocument)]
    internal class Expander : ICommandHandler<TabKeyCommandArgs>
    {
        private static readonly Regex _bracket = new Regex(@"<([a-z0-9]*)\b[^>]*>([^<]*)</\1>", RegexOptions.IgnoreCase);
        private static readonly Regex _quotes = new Regex("(=\"()\")", RegexOptions.IgnoreCase);

        [Import]
        private readonly IEditorCommandHandlerServiceFactory _commandService = default;

        [Import]
        private readonly IClassifierAggregatorService _classifierService = default;

        [Import]
        private readonly ITextBufferUndoManagerProvider _undoProvider = default;

        private static Span _emptySpan = new Span();

        private bool InvokeZenCoding(ITextView view)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Span zenSpan = GetSyntaxSpan(view, out var syntax);

            if (zenSpan.IsEmpty || view.Selection.SelectedSpans[0].Length > 0)
            {
                return false;
            }

            var parser = new Parser();
            var result = parser.Parse(syntax, ZenType.HTML);

            if (!string.IsNullOrEmpty(result))
            {
                ITextBufferUndoManager undoManager = _undoProvider.GetTextBufferUndoManager(view.TextBuffer);

                using (ITextUndoTransaction undo = undoManager.TextBufferUndoHistory.CreateTransaction("ZenCoding"))
                {
                    try
                    {
                        ITextSelection selection = UpdateTextBuffer(view, zenSpan, result);

                        var newSpan = new Span(zenSpan.Start, selection.SelectedSpans[0].Length);

                        IEditorCommandHandlerService service = _commandService.GetService(view);

                        var cmd = new FormatSelectionCommandArgs(view, view.TextBuffer);
                        service.Execute((v, b) => cmd, null);

                        SetCaret(view, newSpan, false);

                        selection.Clear();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.Write(ex);
                    }
                    finally
                    {
                        undo.Complete();
                    }
                }

            }

            return false;
        }


        private bool SetCaret(ITextView view, Span zenSpan, bool isReverse)
        {
            var text = view.TextBuffer.CurrentSnapshot.GetText();
            Span quote = FindTabSpan(zenSpan, isReverse, text, _quotes);
            Span bracket = FindTabSpan(zenSpan, isReverse, text, _bracket);

            if (bracket.Start > 0 && (quote.Start == 0 ||
                                      (!isReverse && (bracket.Start < quote.Start)) ||
                                      (isReverse && (bracket.Start > quote.Start))))
            {
                quote = bracket;
            }

            if (zenSpan.Contains(quote.Start))
            {
                MoveTab(view, quote);
                return true;
            }
            else if (!isReverse)
            {
                MoveTab(view, new Span(zenSpan.End, 0));
                return true;
            }

            return false;
        }

        private void MoveTab(ITextView view, Span quote)
        {
            view.Caret.MoveTo(new SnapshotPoint(view.TextBuffer.CurrentSnapshot, quote.Start));
        }

        private static Span FindTabSpan(Span zenSpan, bool isReverse, string text, Regex regex)
        {
            MatchCollection matches = regex.Matches(text);

            if (!isReverse)
            {
                foreach (Match match in matches)
                {
                    Group group = match.Groups[2];

                    if (group.Index >= zenSpan.Start)
                    {
                        return new Span(group.Index, group.Length);
                    }
                }
            }
            else
            {
                for (var i = matches.Count - 1; i >= 0; i--)
                {
                    Group group = matches[i].Groups[2];

                    if (group.Index < zenSpan.End)
                    {
                        return new Span(group.Index, group.Length);
                    }
                }
            }

            return _emptySpan;
        }

        private ITextSelection UpdateTextBuffer(ITextView view, Span zenSpan, string result)
        {
            view.TextBuffer.Replace(zenSpan, result);

            var point = new SnapshotPoint(view.TextBuffer.CurrentSnapshot, zenSpan.Start);
            var snapshot = new SnapshotSpan(point, result.Length);
            view.Selection.Select(snapshot, false);

            return view.Selection;
        }

        private Span GetSyntaxSpan(ITextView view, out string text)
        {
            var position = view.Caret.Position.BufferPosition.Position;
            text = string.Empty;

            if (position == 0)
            {
                return _emptySpan;
            }

            ITextSnapshotLine line = view.TextBuffer.CurrentSnapshot.GetLineFromPosition(position);
            IClassifier classifier = _classifierService.GetClassifier(view.TextBuffer);
            ClassificationSpan last = classifier.GetClassificationSpans(line.Extent).LastOrDefault();
            SnapshotPoint start = last?.Span.End ?? line.Start;

            if (start > position)
            {
                return _emptySpan;
            }

            text = line.Snapshot.GetText(start, position - start).Trim();
            var offset = position - start - text.Length;

            return Span.FromBounds(start + offset, position);
        }

        public string DisplayName => Vsix.Name;

        public bool ExecuteCommand(TabKeyCommandArgs args, CommandExecutionContext executionContext)
        {
            if (InvokeZenCoding(args.TextView))
            {
                return true;
            }

            return false;
        }

        public CommandState GetCommandState(TabKeyCommandArgs args)
        {
            return CommandState.Available;
        }
    }
}
