using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace ZenCodingVS
{
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("HTML")]
    [ContentType("HTMLX")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    public class CommandRegistration : IVsTextViewCreationListener
    {
        [Import]
        public IVsEditorAdaptersFactoryService EditorAdaptersFactoryService { get; set; }

        [Import]
        public ICompletionBroker CompletionBroker { get; set; }

        [Import]
        private ITextBufferUndoManagerProvider UndoProvider { get; set; }

        [Import]
        private IClassifierAggregatorService ClassifierService { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            var textView = EditorAdaptersFactoryService.GetWpfTextView(textViewAdapter);
            var classifier = ClassifierService.GetClassifier(textView.TextBuffer);

            AddCommandFilter(textViewAdapter, new ExpandCommand(textView, CompletionBroker, UndoProvider, classifier));
        }

        private void AddCommandFilter(IVsTextView textViewAdapter, BaseCommand command)
        {
            textViewAdapter.AddCommandFilter(command, out var next);
            command.Next = next;
        }
    }
}
