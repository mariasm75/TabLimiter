using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace TabLimiter
{

    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(TabLimiterPackage.PackageGuidString)]
    [ProvideOptionPage(typeof(TabLimiterOptions), "Tablimiter", "General", 0, 0, false)]
    public sealed class TabLimiterPackage : AsyncPackage
    {
        /// <summary>
        /// TabLimiterPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "47fc8d2a-c0e2-448f-9404-a52ad69694ed";
        private IVsUIShell _shellService;

        private IVsUIShell7 _shellService7
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return _shellService as IVsUIShell7;
            }
        }

        private uint _cookie;
        private FrameEvents eventListener;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            _shellService = await GetServiceAsync(typeof(IVsUIShell)) as IVsUIShell;
            Assumes.Present(_shellService);

            eventListener = new FrameEvents(GetDialogPage(typeof(TabLimiterOptions)) as TabLimiterOptions);

            if (VSConstants.S_OK == _shellService.GetDocumentWindowEnum(out var frames))
            {
                foreach (IVsWindowFrame frame in new EnumerableWindowsCollection(frames))
                {
                    eventListener.OnFrameCreated(frame);
                }
            }
            _cookie = _shellService7.AdviseWindowFrameEvents(eventListener);
        }

        protected override void Dispose(bool disposing)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            eventListener?.Dispose();
            _shellService7.UnadviseWindowFrameEvents(_cookie);
        }

        private class EnumerableWindowsCollection : EnumerableComCollection<IEnumWindowFrames, IVsWindowFrame>
        {
            public EnumerableWindowsCollection(IEnumWindowFrames windowEnum) : base(windowEnum)
            { }

            public override int Clone(IEnumWindowFrames enumerator, out IEnumWindowFrames clone)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return enumerator.Clone(out clone);
            }

            public override int NextItems(IEnumWindowFrames enumerator, uint count, IVsWindowFrame[] items, out uint fetched)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return enumerator.Next(count, items, out fetched);
            }

            public override int Reset(IEnumWindowFrames enumerator)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return enumerator.Reset();
            }

            public override int Skip(IEnumWindowFrames enumerator, uint count)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return enumerator.Skip(count);
            }
        }
    }

    internal class FrameEvents : IVsWindowFrameEvents, IDisposable
    {
        private IDictionary<IVsWindowFrame, DateTime> _currentWindows = new Dictionary<IVsWindowFrame, DateTime>();
        private TabLimiterOptions _dialogPage;

        public FrameEvents(TabLimiterOptions dialogPage)
        {
            _dialogPage = dialogPage;
            _dialogPage.PropertyChanged += DialogPagePropertyChanged;
        }

        private void DialogPagePropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            KillLastIfTooMuch();
        }

        public void OnFrameCreated(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _currentWindows.Add(frame, DateTime.Now);
            KillLastIfTooMuch();
        }

        public void OnFrameDestroyed(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_currentWindows.ContainsKey(frame))
            {
                _currentWindows.Remove(frame);
            }
            KillLastIfTooMuch();
        }

        public void OnFrameIsVisibleChanged(IVsWindowFrame frame, bool newIsVisible)
        {
        }

        public void OnFrameIsOnScreenChanged(IVsWindowFrame frame, bool newIsOnScreen)
        {
        }

        public void OnActiveFrameChanged(IVsWindowFrame oldFrame, IVsWindowFrame newFrame)
        {
        }

        private void KillLastIfTooMuch()
        {
            var nonPinned = _currentWindows.Where(x => !IsPinned(x.Key)).ToArray();
            if (nonPinned.Length <= _dialogPage.MaxNumberOfTabs)
            {
                return;
            }
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (var window in nonPinned
                .OrderByDescending(x => GetPreviewScore(x.Key))
                .ThenBy(x => x.Value).Select(x => x.Key))
            {
                ErrorHandler.ThrowOnFailure(window.GetProperty((int)__VSFPROPID.VSFPROPID_DocData, out var docData));
                if (docData is IVsPersistDocData pdd && !IsDirty(pdd))
                {
                    window.CloseFrame((uint)__FRAMECLOSE.FRAMECLOSE_NoSave);
                    return;
                }
            }
        }

        private static bool IsDirty(IVsPersistDocData pdd)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (pdd.IsDocDataDirty(out int result) == VSConstants.S_OK)
            {
                return result > 0;
            }
            return true;
        }

        private static bool IsPinned(IVsWindowFrame window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(window.GetProperty((int)__VSFPROPID5.VSFPROPID_IsPinned, out var isPinned));
            return isPinned is bool isp && isp;
        }

        private static int GetPreviewScore(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(frame.GetProperty((int)__VSFPROPID5.VSFPROPID_IsProvisional, out var result));
            if (result == null) { return 0; } else { return (bool)result ? 1 : 0; }
        }

        public void Dispose()
        {
            _dialogPage.PropertyChanged -= DialogPagePropertyChanged;
        }
    }
}