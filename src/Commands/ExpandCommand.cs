using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Projection;
using ZenCoding;

namespace ZenCodingVS
{
    internal class ExpandCommand : BaseCommand
    {
        private static readonly Regex _bracket = new Regex(@"<([a-z0-9]*)\b[^>]*>([^<]*)</\1>", RegexOptions.IgnoreCase);
        private static readonly Regex _quotes = new Regex("(=\"()\")", RegexOptions.IgnoreCase);

        private readonly ICompletionBroker _broker;
        private readonly IWpfTextView _view;
        private readonly ITextBufferUndoManager _undoManager;
        private readonly IClassifier _classifier;
        private static Span _emptySpan = new Span();

        public ExpandCommand(IWpfTextView view, ICompletionBroker broker, ITextBufferUndoManagerProvider undoProvider, IClassifier classifier)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _view = view;
            _broker = broker;
            _undoManager = undoProvider.GetTextBufferUndoManager(view.TextBuffer);
            _classifier = classifier;
        }

        public override int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB && !_broker.IsCompletionActive(_view))
            {
                if (InvokeZenCoding())
                {
                    return VSConstants.S_OK;
                }
            }

            return Next.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        private bool InvokeZenCoding()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Span zenSpan = GetSyntaxSpan(out var syntax);

            if (zenSpan.IsEmpty || _view.Selection.SelectedSpans[0].Length > 0 || !IsValidTextBuffer())
            {
                return false;
            }

            var parser = new Parser();
            var result = parser.Parse(syntax, ZenType.HTML);

            if (!string.IsNullOrEmpty(result))
            {
                using (ITextUndoTransaction undo = _undoManager.TextBufferUndoHistory.CreateTransaction("ZenCoding"))
                {
                    ITextSelection selection = UpdateTextBuffer(zenSpan, result);

                    var newSpan = new Span(zenSpan.Start, selection.SelectedSpans[0].Length);

                    var cmd = new CommandID(new Guid("{1496A755-94DE-11D0-8C3F-00C04FC2AAE2}"), 0x70);
                    cmd.Execute();
                    SetCaret(newSpan, false);

                    selection.Clear();
                    undo.Complete();
                }

                return true;
            }

            return false;
        }

        private bool IsValidTextBuffer()
        {
            if (_view.TextBuffer is IProjectionBuffer projection)
            {
                SnapshotPoint snapshotPoint = _view.Caret.Position.BufferPosition;

                System.Collections.Generic.IEnumerable<ITextBuffer> buffers = projection.SourceBuffers.Where(
                    s =>
                        !s.ContentType.IsOfType("html")
                        && !s.ContentType.IsOfType("htmlx")
                        && !s.ContentType.IsOfType("inert")
                        && !s.ContentType.IsOfType("CSharp")
                        && !s.ContentType.IsOfType("VisualBasic")
                        && !s.ContentType.IsOfType("RoslynCSharp")
                        && !s.ContentType.IsOfType("RoslynVisualBasic"));


                foreach (ITextBuffer buffer in buffers)
                {
                    SnapshotPoint? point = _view.BufferGraph.MapDownToBuffer(snapshotPoint, PointTrackingMode.Negative, buffer, PositionAffinity.Predecessor);

                    if (point.HasValue)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private bool SetCaret(Span zenSpan, bool isReverse)
        {
            var text = _view.TextBuffer.CurrentSnapshot.GetText();
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
                MoveTab(quote);
                return true;
            }
            else if (!isReverse)
            {
                MoveTab(new Span(zenSpan.End, 0));
                return true;
            }

            return false;
        }

        private void MoveTab(Span quote)
        {
            _view.Caret.MoveTo(new SnapshotPoint(_view.TextBuffer.CurrentSnapshot, quote.Start));
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

        private ITextSelection UpdateTextBuffer(Span zenSpan, string result)
        {
            _view.TextBuffer.Replace(zenSpan, result);

            var point = new SnapshotPoint(_view.TextBuffer.CurrentSnapshot, zenSpan.Start);
            var snapshot = new SnapshotSpan(point, result.Length);
            _view.Selection.Select(snapshot, false);

            return _view.Selection;
        }

        private Span GetSyntaxSpan(out string text)
        {
            var position = _view.Caret.Position.BufferPosition.Position;
            text = string.Empty;

            if (position == 0)
            {
                return _emptySpan;
            }

            ITextSnapshotLine line = _view.TextBuffer.CurrentSnapshot.GetLineFromPosition(position);
            ClassificationSpan last = _classifier.GetClassificationSpans(line.Extent).LastOrDefault();
            SnapshotPoint start = last?.Span.End ?? line.Start;

            if (start > position)
            {
                return _emptySpan;
            }

            text = line.Snapshot.GetText(start, position - start).Trim();
            var offset = position - start - text.Length;

            return Span.FromBounds(start + offset, position);
        }

        public override int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (pguidCmdGroup == VSConstants.VSStd2K && prgCmds[0].cmdID == (uint)VSConstants.VSStd2KCmdID.FORMATDOCUMENT)
            {
                prgCmds[0].cmdf = (uint)OLECMDF.OLECMDF_ENABLED | (uint)OLECMDF.OLECMDF_SUPPORTED;
                return VSConstants.S_OK;
            }

            return Next.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }
    }

    public static class CommandExtensions
    {
        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <returns>Returns <see langword="true"/> if the command was succesfully executed; otherwise <see langword="false"/>.</returns>
        public static bool Execute(this CommandID cmd, string argument = "")
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var cs = Package.GetGlobalService(typeof(SUIHostCommandDispatcher)) as IOleCommandTarget;

            var argByteCount = Encoding.Unicode.GetByteCount(argument);
            IntPtr inArgPtr = Marshal.AllocCoTaskMem(argByteCount);

            try
            {
                Marshal.GetNativeVariantForObject(argument, inArgPtr);
                var result = cs.Exec(cmd.Guid, (uint)cmd.ID, (uint)OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT, inArgPtr, IntPtr.Zero);

                return result == VSConstants.S_OK;
            }
            finally
            {
                Marshal.Release(inArgPtr);
            }
        }
    }
}
