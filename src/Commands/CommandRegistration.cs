using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

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
            IWpfTextView textView = EditorAdaptersFactoryService.GetWpfTextView(textViewAdapter);
            IClassifier classifier = ClassifierService.GetClassifier(textView.TextBuffer);

            AddCommandFilter(textViewAdapter, new ExpandCommand(textView, CompletionBroker, UndoProvider, classifier));
        }

        private void AddCommandFilter(IVsTextView textViewAdapter, BaseCommand command)
        {
            textViewAdapter.AddCommandFilter(command, out Microsoft.VisualStudio.OLE.Interop.IOleCommandTarget next);
            command.Next = next;
        }
    }
}
