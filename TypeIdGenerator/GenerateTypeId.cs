using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;

namespace TypeIdGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerateTypeId
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3256532c-6e9a-49ba-aa93-ac78803374ad");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateTypeId"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GenerateTypeId(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerateTypeId Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in GenerateTypeId's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateTypeId(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var service = ServiceProvider.GetServiceAsync(typeof(SVsTextManager)).Result;
            var textManager = service as IVsTextManager2;
            IVsTextView view;
            int res = textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out view);
            if (res != Microsoft.VisualStudio.VSConstants.S_OK)
                return;
            IVsTextLines lines;
            if (view.GetBuffer(out lines) != Microsoft.VisualStudio.VSConstants.S_OK)
                return;
            int line, col;
            if (view.GetCaretPos(out line, out col) != Microsoft.VisualStudio.VSConstants.S_OK)
                return;
            string data;
            int length;
            lines.GetLengthOfLine(line, out length);
            lines.GetLineText(line, 0, line, length, out data);
            TextSpan[] span;

            byte[] valueBits = new byte[4];
            new Random().NextBytes(valueBits);
            UInt32 value = BitConverter.ToUInt32(valueBits, 0);
            if (col > data.Length)
                data = data.PadRight(col);
            data = data.Insert(col, String.Format("static constexpr type_id ID${0,08:x}_id = 0x{0,08:x}U; // {0}", value, value, value));

            var ptr = IntPtr.Zero;

            try
            {
                ptr = Marshal.StringToHGlobalAuto(data);
                lines.ReplaceLines(line, 0, line, length, ptr, data.Length, new[] { new TextSpan() });
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                    Marshal.FreeHGlobal(ptr);
            }
        }
    }
}
