using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelTools
{
    public class ExcelApplication : Excel.Application, IDisposable
    {
        public ExcelApplication()
        {
            Application = new Excel.Application();
            if (Application == null)
            {
                throw new Exception("EXCEL could not be started. Check that your office installation and project references are correct.");
            }
        }

        public Excel.Application Application { get; private set; }

        public void Dispose()
        {
            Application.Quit();
            try
            {
                int result;
                do
                {
                    result = Application.ReleaseObject();
                }
                while (result > 0);
            }
            catch (Exception ex)
            {
                Application = null;
                throw new Exception("Error releasing object: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
            Application = null;
        }

        #region Excel.Application
        public Excel.Range ActiveCell
        {
            get
            {
                return Application.ActiveCell;
            }
        }

        public Excel.Chart ActiveChart
        {
            get
            {
                return Application.ActiveChart;
            }
        }

        public Excel.DialogSheet ActiveDialog
        {
            get
            {
                return Application.ActiveDialog;
            }
        }

        public int ActiveEncryptionSession
        {
            get
            {
                return Application.ActiveEncryptionSession;
            }
        }

        public Excel.MenuBar ActiveMenuBar
        {
            get
            {
                return Application.ActiveMenuBar;
            }
        }

        public string ActivePrinter
        {
            get
            {
                return Application.ActivePrinter;
            }

            set
            {
                Application.ActivePrinter = value;
            }
        }

        public Excel.ProtectedViewWindow ActiveProtectedViewWindow
        {
            get
            {
                return Application.ActiveProtectedViewWindow;
            }
        }

        public dynamic ActiveSheet
        {
            get
            {
                return Application.ActiveSheet;
            }
        }

        public Excel.Window ActiveWindow
        {
            get
            {
                return Application.ActiveWindow;
            }
        }

        public Excel.Workbook ActiveWorkbook
        {
            get
            {
                return Application.ActiveWorkbook;
            }
        }

        public Excel.AddIns AddIns
        {
            get
            {
                return Application.AddIns;
            }
        }

        public Excel.AddIns2 AddIns2
        {
            get
            {
                return Application.AddIns2;
            }
        }

        public bool AlertBeforeOverwriting
        {
            get
            {
                return Application.AlertBeforeOverwriting;
            }

            set
            {
                Application.AlertBeforeOverwriting = value;
            }
        }

        public string AltStartupPath
        {
            get
            {
                return Application.AltStartupPath;
            }

            set
            {
                Application.AltStartupPath = value;
            }
        }

        public bool AlwaysUseClearType
        {
            get
            {
                return Application.AlwaysUseClearType;
            }

            set
            {
                Application.AlwaysUseClearType = value;
            }
        }

        public AnswerWizard AnswerWizard
        {
            get
            {
                return Application.AnswerWizard;
            }
        }

        public bool ArbitraryXMLSupportAvailable
        {
            get
            {
                return Application.ArbitraryXMLSupportAvailable;
            }
        }

        public bool AskToUpdateLinks
        {
            get
            {
                return Application.AskToUpdateLinks;
            }

            set
            {
                Application.AskToUpdateLinks = value;
            }
        }

        public IAssistance Assistance
        {
            get
            {
                return Application.Assistance;
            }
        }

        public Assistant Assistant
        {
            get
            {
                return Application.Assistant;
            }
        }

        public Excel.AutoCorrect AutoCorrect
        {
            get
            {
                return Application.AutoCorrect;
            }
        }

        public bool AutoFormatAsYouTypeReplaceHyperlinks
        {
            get
            {
                return Application.AutoFormatAsYouTypeReplaceHyperlinks;
            }

            set
            {
                Application.AutoFormatAsYouTypeReplaceHyperlinks = value;
            }
        }

        public MsoAutomationSecurity AutomationSecurity
        {
            get
            {
                return Application.AutomationSecurity;
            }

            set
            {
                Application.AutomationSecurity = value;
            }
        }

        public bool AutoPercentEntry
        {
            get
            {
                return Application.AutoPercentEntry;
            }

            set
            {
                Application.AutoPercentEntry = value;
            }
        }

        public Excel.AutoRecover AutoRecover
        {
            get
            {
                return Application.AutoRecover;
            }
        }

        public int Build
        {
            get
            {
                return Application.Build;
            }
        }

        public bool CalculateBeforeSave
        {
            get
            {
                return Application.CalculateBeforeSave;
            }

            set
            {
                Application.CalculateBeforeSave = value;
            }
        }

        public Excel.XlCalculation Calculation
        {
            get
            {
                return Application.Calculation;
            }

            set
            {
                Application.Calculation = value;
            }
        }

        public Excel.XlCalculationInterruptKey CalculationInterruptKey
        {
            get
            {
                return Application.CalculationInterruptKey;
            }

            set
            {
                Application.CalculationInterruptKey = value;
            }
        }

        public Excel.XlCalculationState CalculationState
        {
            get
            {
                return Application.CalculationState;
            }
        }

        public int CalculationVersion
        {
            get
            {
                return Application.CalculationVersion;
            }
        }

        public bool CanPlaySounds
        {
            get
            {
                return Application.CanPlaySounds;
            }
        }

        public bool CanRecordSounds
        {
            get
            {
                return Application.CanRecordSounds;
            }
        }

        public string Caption
        {
            get
            {
                return Application.Caption;
            }

            set
            {
                Application.Caption = value;
            }
        }

        public bool CellDragAndDrop
        {
            get
            {
                return Application.CellDragAndDrop;
            }

            set
            {
                Application.CellDragAndDrop = value;
            }
        }

        public Excel.Range Cells
        {
            get
            {
                return Application.Cells;
            }
        }

        public Excel.Sheets Charts
        {
            get
            {
                return Application.Charts;
            }
        }

        public string ClusterConnector
        {
            get
            {
                return Application.ClusterConnector;
            }

            set
            {
                Application.ClusterConnector = value;
            }
        }

        public bool ColorButtons
        {
            get
            {
                return Application.ColorButtons;
            }

            set
            {
                Application.ColorButtons = value;
            }
        }

        public Excel.Range Columns
        {
            get
            {
                return Application.Columns;
            }
        }

        public COMAddIns COMAddIns
        {
            get
            {
                return Application.COMAddIns;
            }
        }

        public CommandBars CommandBars
        {
            get
            {
                return Application.CommandBars;
            }
        }

        public Excel.XlCommandUnderlines CommandUnderlines
        {
            get
            {
                return Application.CommandUnderlines;
            }

            set
            {
                Application.CommandUnderlines = value;
            }
        }

        public bool ConstrainNumeric
        {
            get
            {
                return Application.ConstrainNumeric;
            }

            set
            {
                Application.ConstrainNumeric = value;
            }
        }

        public bool ControlCharacters
        {
            get
            {
                return Application.ControlCharacters;
            }

            set
            {
                Application.ControlCharacters = value;
            }
        }

        public bool CopyObjectsWithCells
        {
            get
            {
                return Application.CopyObjectsWithCells;
            }

            set
            {
                Application.CopyObjectsWithCells = value;
            }
        }

        public Excel.XlCreator Creator
        {
            get
            {
                return Application.Creator;
            }
        }

        public Excel.XlMousePointer Cursor
        {
            get
            {
                return Application.Cursor;
            }

            set
            {
                Application.Cursor = value;
            }
        }

        public int CursorMovement
        {
            get
            {
                return Application.CursorMovement;
            }

            set
            {
                Application.CursorMovement = value;
            }
        }

        public int CustomListCount
        {
            get
            {
                return Application.CustomListCount;
            }
        }

        public Excel.XlCutCopyMode CutCopyMode
        {
            get
            {
                return Application.CutCopyMode;
            }

            set
            {
                Application.CutCopyMode = value;
            }
        }

        public int DataEntryMode
        {
            get
            {
                return Application.DataEntryMode;
            }

            set
            {
                Application.DataEntryMode = value;
            }
        }

        public int DDEAppReturnCode
        {
            get
            {
                return Application.DDEAppReturnCode;
            }
        }

        public string DecimalSeparator
        {
            get
            {
                return Application.DecimalSeparator;
            }

            set
            {
                Application.DecimalSeparator = value;
            }
        }

        public string DefaultFilePath
        {
            get
            {
                return Application.DefaultFilePath;
            }

            set
            {
                Application.DefaultFilePath = value;
            }
        }

        public Excel.XlFileFormat DefaultSaveFormat
        {
            get
            {
                return Application.DefaultSaveFormat;
            }

            set
            {
                Application.DefaultSaveFormat = value;
            }
        }

        public int DefaultSheetDirection
        {
            get
            {
                return Application.DefaultSheetDirection;
            }

            set
            {
                Application.DefaultSheetDirection = value;
            }
        }

        public Excel.DefaultWebOptions DefaultWebOptions
        {
            get
            {
                return Application.DefaultWebOptions;
            }
        }

        public bool DeferAsyncQueries
        {
            get
            {
                return Application.DeferAsyncQueries;
            }

            set
            {
                Application.DeferAsyncQueries = value;
            }
        }

        public Excel.Dialogs Dialogs
        {
            get
            {
                return Application.Dialogs;
            }
        }

        public Excel.Sheets DialogSheets
        {
            get
            {
                return Application.DialogSheets;
            }
        }

        public bool DisplayAlerts
        {
            get
            {
                return Application.DisplayAlerts;
            }

            set
            {
                Application.DisplayAlerts = value;
            }
        }

        public bool DisplayClipboardWindow
        {
            get
            {
                return Application.DisplayClipboardWindow;
            }

            set
            {
                Application.DisplayClipboardWindow = value;
            }
        }

        public Excel.XlCommentDisplayMode DisplayCommentIndicator
        {
            get
            {
                return Application.DisplayCommentIndicator;
            }

            set
            {
                Application.DisplayCommentIndicator = value;
            }
        }

        public bool DisplayDocumentActionTaskPane
        {
            get
            {
                return Application.DisplayDocumentActionTaskPane;
            }

            set
            {
                Application.DisplayDocumentActionTaskPane = value;
            }
        }

        public bool DisplayDocumentInformationPanel
        {
            get
            {
                return Application.DisplayDocumentInformationPanel;
            }

            set
            {
                Application.DisplayDocumentInformationPanel = value;
            }
        }

        public bool DisplayExcel4Menus
        {
            get
            {
                return Application.DisplayExcel4Menus;
            }

            set
            {
                Application.DisplayExcel4Menus = value;
            }
        }

        public bool DisplayFormulaAutoComplete
        {
            get
            {
                return Application.DisplayFormulaAutoComplete;
            }

            set
            {
                Application.DisplayFormulaAutoComplete = value;
            }
        }

        public bool DisplayFormulaBar
        {
            get
            {
                return Application.DisplayFormulaBar;
            }

            set
            {
                Application.DisplayFormulaBar = value;
            }
        }

        public bool DisplayFullScreen
        {
            get
            {
                return Application.DisplayFullScreen;
            }

            set
            {
                Application.DisplayFullScreen = value;
            }
        }

        public bool DisplayFunctionToolTips
        {
            get
            {
                return Application.DisplayFunctionToolTips;
            }

            set
            {
                Application.DisplayFunctionToolTips = value;
            }
        }

        public bool DisplayInfoWindow
        {
            get
            {
                return Application.DisplayInfoWindow;
            }

            set
            {
                Application.DisplayInfoWindow = value;
            }
        }

        public bool DisplayInsertOptions
        {
            get
            {
                return Application.DisplayInsertOptions;
            }

            set
            {
                Application.DisplayInsertOptions = value;
            }
        }

        public bool DisplayNoteIndicator
        {
            get
            {
                return Application.DisplayNoteIndicator;
            }

            set
            {
                Application.DisplayNoteIndicator = value;
            }
        }

        public bool DisplayPasteOptions
        {
            get
            {
                return Application.DisplayPasteOptions;
            }

            set
            {
                Application.DisplayPasteOptions = value;
            }
        }

        public bool DisplayRecentFiles
        {
            get
            {
                return Application.DisplayRecentFiles;
            }

            set
            {
                Application.DisplayRecentFiles = value;
            }
        }

        public bool DisplayScrollBars
        {
            get
            {
                return Application.DisplayScrollBars;
            }

            set
            {
                Application.DisplayScrollBars = value;
            }
        }

        public bool DisplayStatusBar
        {
            get
            {
                return Application.DisplayStatusBar;
            }

            set
            {
                Application.DisplayStatusBar = value;
            }
        }

        public dynamic Dummy101
        {
            get
            {
                return Application.Dummy101;
            }
        }

        public bool Dummy22
        {
            get
            {
                return Application.Dummy22;
            }

            set
            {
                Application.Dummy22 = value;
            }
        }

        public bool Dummy23
        {
            get
            {
                return Application.Dummy23;
            }

            set
            {
                Application.Dummy23 = value;
            }
        }

        public bool EditDirectlyInCell
        {
            get
            {
                return Application.EditDirectlyInCell;
            }

            set
            {
                Application.EditDirectlyInCell = value;
            }
        }

        public bool EnableAnimations
        {
            get
            {
                return Application.EnableAnimations;
            }

            set
            {
                Application.EnableAnimations = value;
            }
        }

        public bool EnableAutoComplete
        {
            get
            {
                return Application.EnableAutoComplete;
            }

            set
            {
                Application.EnableAutoComplete = value;
            }
        }

        public Excel.XlEnableCancelKey EnableCancelKey
        {
            get
            {
                return Application.EnableCancelKey;
            }

            set
            {
                Application.EnableCancelKey = value;
            }
        }

        public bool EnableEvents
        {
            get
            {
                return Application.EnableEvents;
            }

            set
            {
                Application.EnableEvents = value;
            }
        }

        public bool EnableLargeOperationAlert
        {
            get
            {
                return Application.EnableLargeOperationAlert;
            }

            set
            {
                Application.EnableLargeOperationAlert = value;
            }
        }

        public bool EnableLivePreview
        {
            get
            {
                return Application.EnableLivePreview;
            }

            set
            {
                Application.EnableLivePreview = value;
            }
        }

        public bool EnableSound
        {
            get
            {
                return Application.EnableSound;
            }

            set
            {
                Application.EnableSound = value;
            }
        }

        public bool EnableTipWizard
        {
            get
            {
                return Application.EnableTipWizard;
            }

            set
            {
                Application.EnableTipWizard = value;
            }
        }

        public Excel.ErrorCheckingOptions ErrorCheckingOptions
        {
            get
            {
                return Application.ErrorCheckingOptions;
            }
        }

        public Excel.Sheets Excel4IntlMacroSheets
        {
            get
            {
                return Application.Excel4IntlMacroSheets;
            }
        }

        public Excel.Sheets Excel4MacroSheets
        {
            get
            {
                return Application.Excel4MacroSheets;
            }
        }

        public bool ExtendList
        {
            get
            {
                return Application.ExtendList;
            }

            set
            {
                Application.ExtendList = value;
            }
        }

        public MsoFeatureInstall FeatureInstall
        {
            get
            {
                return Application.FeatureInstall;
            }

            set
            {
                Application.FeatureInstall = value;
            }
        }

        public Excel.FileExportConverters FileExportConverters
        {
            get
            {
                return Application.FileExportConverters;
            }
        }

        public IFind FileFind
        {
            get
            {
                return Application.FileFind;
            }
        }

        public FileSearch FileSearch
        {
            get
            {
                return Application.FileSearch;
            }
        }

        public MsoFileValidationMode FileValidation
        {
            get
            {
                return Application.FileValidation;
            }

            set
            {
                Application.FileValidation = value;
            }
        }

        public Excel.XlFileValidationPivotMode FileValidationPivot
        {
            get
            {
                return Application.FileValidationPivot;
            }

            set
            {
                Application.FileValidationPivot = value;
            }
        }

        public Excel.CellFormat FindFormat
        {
            get
            {
                return Application.FindFormat;
            }

            set
            {
                Application.FindFormat = value;
            }
        }

        public bool FixedDecimal
        {
            get
            {
                return Application.FixedDecimal;
            }

            set
            {
                Application.FixedDecimal = value;
            }
        }

        public int FixedDecimalPlaces
        {
            get
            {
                return Application.FixedDecimalPlaces;
            }

            set
            {
                Application.FixedDecimalPlaces = value;
            }
        }

        public int FormulaBarHeight
        {
            get
            {
                return Application.FormulaBarHeight;
            }

            set
            {
                Application.FormulaBarHeight = value;
            }
        }

        public bool GenerateGetPivotData
        {
            get
            {
                return Application.GenerateGetPivotData;
            }

            set
            {
                Application.GenerateGetPivotData = value;
            }
        }

        public Excel.XlGenerateTableRefs GenerateTableRefs
        {
            get
            {
                return Application.GenerateTableRefs;
            }

            set
            {
                Application.GenerateTableRefs = value;
            }
        }

        public double Height
        {
            get
            {
                return Application.Height;
            }

            set
            {
                Application.Height = value;
            }
        }

        public bool HighQualityModeForGraphics
        {
            get
            {
                return Application.HighQualityModeForGraphics;
            }

            set
            {
                Application.HighQualityModeForGraphics = value;
            }
        }

        public int Hinstance
        {
            get
            {
                return Application.Hinstance;
            }
        }

        public dynamic HinstancePtr
        {
            get
            {
                return Application.HinstancePtr;
            }
        }

        public int Hwnd
        {
            get
            {
                return Application.Hwnd;
            }
        }

        public bool IgnoreRemoteRequests
        {
            get
            {
                return Application.IgnoreRemoteRequests;
            }

            set
            {
                Application.IgnoreRemoteRequests = value;
            }
        }

        public bool Interactive
        {
            get
            {
                return Application.Interactive;
            }

            set
            {
                Application.Interactive = value;
            }
        }

        public bool IsSandboxed
        {
            get
            {
                return Application.IsSandboxed;
            }
        }

        public bool Iteration
        {
            get
            {
                return Application.Iteration;
            }

            set
            {
                Application.Iteration = value;
            }
        }

        public LanguageSettings LanguageSettings
        {
            get
            {
                return Application.LanguageSettings;
            }
        }

        public bool LargeButtons
        {
            get
            {
                return Application.LargeButtons;
            }

            set
            {
                Application.LargeButtons = value;
            }
        }

        public int LargeOperationCellThousandCount
        {
            get
            {
                return Application.LargeOperationCellThousandCount;
            }

            set
            {
                Application.LargeOperationCellThousandCount = value;
            }
        }

        public double Left
        {
            get
            {
                return Application.Left;
            }

            set
            {
                Application.Left = value;
            }
        }

        public string LibraryPath
        {
            get
            {
                return Application.LibraryPath;
            }
        }

        public dynamic MailSession
        {
            get
            {
                return Application.MailSession;
            }
        }

        public Excel.XlMailSystem MailSystem
        {
            get
            {
                return Application.MailSystem;
            }
        }

        public bool MapPaperSize
        {
            get
            {
                return Application.MapPaperSize;
            }

            set
            {
                Application.MapPaperSize = value;
            }
        }

        public bool MathCoprocessorAvailable
        {
            get
            {
                return Application.MathCoprocessorAvailable;
            }
        }

        public double MaxChange
        {
            get
            {
                return Application.MaxChange;
            }

            set
            {
                Application.MaxChange = value;
            }
        }

        public int MaxIterations
        {
            get
            {
                return Application.MaxIterations;
            }

            set
            {
                Application.MaxIterations = value;
            }
        }

        public int MeasurementUnit
        {
            get
            {
                return Application.MeasurementUnit;
            }

            set
            {
                Application.MeasurementUnit = value;
            }
        }

        public int MemoryFree
        {
            get
            {
                return Application.MemoryFree;
            }
        }

        public int MemoryTotal
        {
            get
            {
                return Application.MemoryTotal;
            }
        }

        public int MemoryUsed
        {
            get
            {
                return Application.MemoryUsed;
            }
        }

        public Excel.MenuBars MenuBars
        {
            get
            {
                return Application.MenuBars;
            }
        }

        public Excel.Modules Modules
        {
            get
            {
                return Application.Modules;
            }
        }

        public bool MouseAvailable
        {
            get
            {
                return Application.MouseAvailable;
            }
        }

        public bool MoveAfterReturn
        {
            get
            {
                return Application.MoveAfterReturn;
            }

            set
            {
                Application.MoveAfterReturn = value;
            }
        }

        public Excel.XlDirection MoveAfterReturnDirection
        {
            get
            {
                return Application.MoveAfterReturnDirection;
            }

            set
            {
                Application.MoveAfterReturnDirection = value;
            }
        }

        public Excel.MultiThreadedCalculation MultiThreadedCalculation
        {
            get
            {
                return Application.MultiThreadedCalculation;
            }
        }

        public string Name
        {
            get
            {
                return Application.Name;
            }
        }

        public Excel.Names Names
        {
            get
            {
                return Application.Names;
            }
        }

        public string NetworkTemplatesPath
        {
            get
            {
                return Application.NetworkTemplatesPath;
            }
        }

        public NewFile NewWorkbook
        {
            get
            {
                throw new NotImplementedException();
              //  return Application.NewWorkbook;
            }
        }

        public Excel.ODBCErrors ODBCErrors
        {
            get
            {
                return Application.ODBCErrors;
            }
        }

        public int ODBCTimeout
        {
            get
            {
                return Application.ODBCTimeout;
            }

            set
            {
                Application.ODBCTimeout = value;
            }
        }

        public Excel.OLEDBErrors OLEDBErrors
        {
            get
            {
                return Application.OLEDBErrors;
            }
        }

        public string OnCalculate
        {
            get
            {
                return Application.OnCalculate;
            }

            set
            {
                Application.OnCalculate = value;
            }
        }

        public string OnData
        {
            get
            {
                return Application.OnData;
            }

            set
            {
                Application.OnData = value;
            }
        }

        public string OnDoubleClick
        {
            get
            {
                return Application.OnDoubleClick;
            }

            set
            {
                Application.OnDoubleClick = value;
            }
        }

        public string OnEntry
        {
            get
            {
                return Application.OnEntry;
            }

            set
            {
                Application.OnEntry = value;
            }
        }

        public string OnSheetActivate
        {
            get
            {
                return Application.OnSheetActivate;
            }

            set
            {
                Application.OnSheetActivate = value;
            }
        }

        public string OnSheetDeactivate
        {
            get
            {
                return Application.OnSheetDeactivate;
            }

            set
            {
                Application.OnSheetDeactivate = value;
            }
        }

        public string OnWindow
        {
            get
            {
                return Application.OnWindow;
            }

            set
            {
                Application.OnWindow = value;
            }
        }

        public string OperatingSystem
        {
            get
            {
                return Application.OperatingSystem;
            }
        }

        public string OrganizationName
        {
            get
            {
                return Application.OrganizationName;
            }
        }

        public Excel.Application Parent
        {
            get
            {
                return Application.Parent;
            }
        }

        public string Path
        {
            get
            {
                return Application.Path;
            }
        }

        public string PathSeparator
        {
            get
            {
                return Application.PathSeparator;
            }
        }

        public bool PivotTableSelection
        {
            get
            {
                return Application.PivotTableSelection;
            }

            set
            {
                Application.PivotTableSelection = value;
            }
        }

        public bool PrintCommunication
        {
            get
            {
                return Application.PrintCommunication;
            }

            set
            {
                Application.PrintCommunication = value;
            }
        }

        public string ProductCode
        {
            get
            {
                return Application.ProductCode;
            }
        }

        public bool PromptForSummaryInfo
        {
            get
            {
                return Application.PromptForSummaryInfo;
            }

            set
            {
                Application.PromptForSummaryInfo = value;
            }
        }

        public Excel.ProtectedViewWindows ProtectedViewWindows
        {
            get
            {
                return Application.ProtectedViewWindows;
            }
        }

        public bool Quitting
        {
            get
            {
                return Application.Quitting;
            }
        }

        public bool Ready
        {
            get
            {
                return Application.Ready;
            }
        }

        public Excel.RecentFiles RecentFiles
        {
            get
            {
                return Application.RecentFiles;
            }
        }

        public bool RecordRelative
        {
            get
            {
                return Application.RecordRelative;
            }
        }

        public Excel.XlReferenceStyle ReferenceStyle
        {
            get
            {
                return Application.ReferenceStyle;
            }

            set
            {
                Application.ReferenceStyle = value;
            }
        }

        public Excel.CellFormat ReplaceFormat
        {
            get
            {
                return Application.ReplaceFormat;
            }

            set
            {
                Application.ReplaceFormat = value;
            }
        }

        public bool RollZoom
        {
            get
            {
                return Application.RollZoom;
            }

            set
            {
                Application.RollZoom = value;
            }
        }

        public Excel.Range Rows
        {
            get
            {
                return Application.Rows;
            }
        }

        public Excel.RTD RTD
        {
            get
            {
                return Application.RTD;
            }
        }

        public bool SaveISO8601Dates
        {
            get
            {
                return Application.SaveISO8601Dates;
            }

            set
            {
                Application.SaveISO8601Dates = value;
            }
        }

        public bool ScreenUpdating
        {
            get
            {
                return Application.ScreenUpdating;
            }

            set
            {
                Application.ScreenUpdating = value;
            }
        }

        public dynamic Selection
        {
            get
            {
                return Application.Selection;
            }
        }

        public Excel.Sheets Sheets
        {
            get
            {
                return Application.Sheets;
            }
        }

        public int SheetsInNewWorkbook
        {
            get
            {
                return Application.SheetsInNewWorkbook;
            }

            set
            {
                Application.SheetsInNewWorkbook = value;
            }
        }

        public bool ShowChartTipNames
        {
            get
            {
                return Application.ShowChartTipNames;
            }

            set
            {
                Application.ShowChartTipNames = value;
            }
        }

        public bool ShowChartTipValues
        {
            get
            {
                return Application.ShowChartTipValues;
            }

            set
            {
                Application.ShowChartTipValues = value;
            }
        }

        public bool ShowDevTools
        {
            get
            {
                return Application.ShowDevTools;
            }

            set
            {
                Application.ShowDevTools = value;
            }
        }

        public bool ShowMenuFloaties
        {
            get
            {
                return Application.ShowMenuFloaties;
            }

            set
            {
                Application.ShowMenuFloaties = value;
            }
        }

        public bool ShowSelectionFloaties
        {
            get
            {
                return Application.ShowSelectionFloaties;
            }

            set
            {
                Application.ShowSelectionFloaties = value;
            }
        }

        public bool ShowStartupDialog
        {
            get
            {
                return Application.ShowStartupDialog;
            }

            set
            {
                Application.ShowStartupDialog = value;
            }
        }

        public bool ShowToolTips
        {
            get
            {
                return Application.ShowToolTips;
            }

            set
            {
                Application.ShowToolTips = value;
            }
        }

        public bool ShowWindowsInTaskbar
        {
            get
            {
                return Application.ShowWindowsInTaskbar;
            }

            set
            {
                Application.ShowWindowsInTaskbar = value;
            }
        }

        public SmartArtColors SmartArtColors
        {
            get
            {
                return Application.SmartArtColors;
            }
        }

        public SmartArtLayouts SmartArtLayouts
        {
            get
            {
                return Application.SmartArtLayouts;
            }
        }

        public SmartArtQuickStyles SmartArtQuickStyles
        {
            get
            {
                return Application.SmartArtQuickStyles;
            }
        }

        public Excel.SmartTagRecognizers SmartTagRecognizers
        {
            get
            {
                return Application.SmartTagRecognizers;
            }
        }

        public Excel.Speech Speech
        {
            get
            {
                return Application.Speech;
            }
        }

        public Excel.SpellingOptions SpellingOptions
        {
            get
            {
                return Application.SpellingOptions;
            }
        }

        public string StandardFont
        {
            get
            {
                return Application.StandardFont;
            }

            set
            {
                Application.StandardFont = value;
            }
        }

        public double StandardFontSize
        {
            get
            {
                return Application.StandardFontSize;
            }

            set
            {
                Application.StandardFontSize = value;
            }
        }

        public string StartupPath
        {
            get
            {
                return Application.StartupPath;
            }
        }

        public dynamic StatusBar
        {
            get
            {
                return Application.StatusBar;
            }

            set
            {
                Application.StatusBar = value;
            }
        }

        public string TemplatesPath
        {
            get
            {
                return Application.TemplatesPath;
            }
        }

        public Excel.Range ThisCell
        {
            get
            {
                return Application.ThisCell;
            }
        }

        public Excel.Workbook ThisWorkbook
        {
            get
            {
                return Application.ThisWorkbook;
            }
        }

        public string ThousandsSeparator
        {
            get
            {
                return Application.ThousandsSeparator;
            }

            set
            {
                Application.ThousandsSeparator = value;
            }
        }

        public Excel.Toolbars Toolbars
        {
            get
            {
                return Application.Toolbars;
            }
        }

        public double Top
        {
            get
            {
                return Application.Top;
            }

            set
            {
                Application.Top = value;
            }
        }

        public string TransitionMenuKey
        {
            get
            {
                return Application.TransitionMenuKey;
            }

            set
            {
                Application.TransitionMenuKey = value;
            }
        }

        public int TransitionMenuKeyAction
        {
            get
            {
                return Application.TransitionMenuKeyAction;
            }

            set
            {
                Application.TransitionMenuKeyAction = value;
            }
        }

        public bool TransitionNavigKeys
        {
            get
            {
                return Application.TransitionNavigKeys;
            }

            set
            {
                Application.TransitionNavigKeys = value;
            }
        }

        public int UILanguage
        {
            get
            {
                return Application.UILanguage;
            }

            set
            {
                Application.UILanguage = value;
            }
        }

        public double UsableHeight
        {
            get
            {
                return Application.UsableHeight;
            }
        }

        public double UsableWidth
        {
            get
            {
                return Application.UsableWidth;
            }
        }

        public bool UseClusterConnector
        {
            get
            {
                return Application.UseClusterConnector;
            }

            set
            {
                Application.UseClusterConnector = value;
            }
        }

        public Excel.UsedObjects UsedObjects
        {
            get
            {
                return Application.UsedObjects;
            }
        }

        public bool UserControl
        {
            get
            {
                return Application.UserControl;
            }

            set
            {
                Application.UserControl = value;
            }
        }

        public string UserLibraryPath
        {
            get
            {
                return Application.UserLibraryPath;
            }
        }

        public string UserName
        {
            get
            {
                return Application.UserName;
            }

            set
            {
                Application.UserName = value;
            }
        }

        public bool UseSystemSeparators
        {
            get
            {
                return Application.UseSystemSeparators;
            }

            set
            {
                Application.UseSystemSeparators = value;
            }
        }

        public string Value
        {
            get
            {
                return Application.Value;
            }
        }

        public VBE VBE
        {
            get
            {
                return Application.VBE;
            }
        }

        public string Version
        {
            get
            {
                return Application.Version;
            }
        }

        public bool Visible
        {
            get
            {
                return Application.Visible;
            }

            set
            {
                Application.Visible = value;
            }
        }

        public bool WarnOnFunctionNameConflict
        {
            get
            {
                return Application.WarnOnFunctionNameConflict;
            }

            set
            {
                Application.WarnOnFunctionNameConflict = value;
            }
        }

        public Excel.Watches Watches
        {
            get
            {
                return Application.Watches;
            }
        }

        public double Width
        {
            get
            {
                return Application.Width;
            }

            set
            {
                Application.Width = value;
            }
        }

        public Excel.Windows Windows
        {
            get
            {
                return Application.Windows;
            }
        }

        public bool WindowsForPens
        {
            get
            {
                return Application.WindowsForPens;
            }
        }

        public Excel.XlWindowState WindowState
        {
            get
            {
                return Application.WindowState;
            }

            set
            {
                Application.WindowState = value;
            }
        }

        public Excel.Workbooks Workbooks
        {
            get
            {
                return Application.Workbooks;
            }
        }

        public Excel.WorksheetFunction WorksheetFunction
        {
            get
            {
                return Application.WorksheetFunction;
            }
        }

        public Excel.Sheets Worksheets
        {
            get
            {
                return Application.Worksheets;
            }
        }

        public string _Default
        {
            get
            {
                return Application._Default;
            }
        }

        public event Excel.AppEvents_AfterCalculateEventHandler AfterCalculate;
        public event Excel.AppEvents_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivate;
        public event Excel.AppEvents_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeClose;
        public event Excel.AppEvents_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEdit;
        public event Excel.AppEvents_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivate;
        public event Excel.AppEvents_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpen;
        public event Excel.AppEvents_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResize;
        public event Excel.AppEvents_SheetActivateEventHandler SheetActivate;
        public event Excel.AppEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        public event Excel.AppEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        public event Excel.AppEvents_SheetCalculateEventHandler SheetCalculate;
        public event Excel.AppEvents_SheetChangeEventHandler SheetChange;
        public event Excel.AppEvents_SheetDeactivateEventHandler SheetDeactivate;
        public event Excel.AppEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        public event Excel.AppEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChange;
        public event Excel.AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChanges;
        public event Excel.AppEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChanges;
        public event Excel.AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChanges;
        public event Excel.AppEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        public event Excel.AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        public event Excel.AppEvents_WindowActivateEventHandler WindowActivate;
        public event Excel.AppEvents_WindowDeactivateEventHandler WindowDeactivate;
        public event Excel.AppEvents_WindowResizeEventHandler WindowResize;
        public event Excel.AppEvents_WorkbookActivateEventHandler WorkbookActivate;
        public event Excel.AppEvents_WorkbookAddinInstallEventHandler WorkbookAddinInstall;
        public event Excel.AppEvents_WorkbookAddinUninstallEventHandler WorkbookAddinUninstall;
        public event Excel.AppEvents_WorkbookAfterSaveEventHandler WorkbookAfterSave;
        public event Excel.AppEvents_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExport;
        public event Excel.AppEvents_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImport;
        public event Excel.AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeClose;
        public event Excel.AppEvents_WorkbookBeforePrintEventHandler WorkbookBeforePrint;
        public event Excel.AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSave;
        public event Excel.AppEvents_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExport;
        public event Excel.AppEvents_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImport;
        public event Excel.AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivate;
        public event Excel.AppEvents_WorkbookNewChartEventHandler WorkbookNewChart;
        public event Excel.AppEvents_WorkbookNewSheetEventHandler WorkbookNewSheet;
        public event Excel.AppEvents_WorkbookOpenEventHandler WorkbookOpen;
        public event Excel.AppEvents_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnection;
        public event Excel.AppEvents_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnection;
        public event Excel.AppEvents_WorkbookRowsetCompleteEventHandler WorkbookRowsetComplete;
        public event Excel.AppEvents_WorkbookSyncEventHandler WorkbookSync;

        event Excel.AppEvents_NewWorkbookEventHandler Excel.AppEvents_Event.NewWorkbook
        {
            add
            {
                throw new NotImplementedException();
            }

            remove
            {
                throw new NotImplementedException();
            }
        }

        public void ActivateMicrosoftApp(Excel.XlMSApplication Index)
        {
            Application.ActivateMicrosoftApp(Index);
        }

        public void AddChartAutoFormat(object Chart, string Name, object Description)
        {
            Application.AddChartAutoFormat(Chart, Name, Description);
        }

        public void AddCustomList(object ListArray, object ByRow)
        {
            Application.AddCustomList(ListArray, ByRow);
        }

        public void Calculate()
        {
            Application.Calculate();
        }

        public void CalculateFull()
        {
            Application.CalculateFull();
        }

        public void CalculateFullRebuild()
        {
            Application.CalculateFullRebuild();
        }

        public void CalculateUntilAsyncQueriesDone()
        {
            Application.CalculateUntilAsyncQueriesDone();
        }

        public double CentimetersToPoints(double Centimeters)
        {
            return Application.CentimetersToPoints(Centimeters);
        }

        public void CheckAbort(object KeepAbort)
        {
            Application.CheckAbort(KeepAbort);
        }

        public bool CheckSpelling(string Word, object CustomDictionary, object IgnoreUppercase)
        {
            return Application.CheckSpelling(Word, CustomDictionary, IgnoreUppercase);
        }

        public dynamic ConvertFormula(object Formula, Excel.XlReferenceStyle FromReferenceStyle, object ToReferenceStyle, object ToAbsolute, object RelativeTo)
        {
            return this.Application.ConvertFormula(Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo);
        }

        public void DDEExecute(int Channel, string String)
        {
            Application.DDEExecute(Channel, String);
        }

        public int DDEInitiate(string App, string Topic)
        {
            return Application.DDEInitiate(App, Topic);
        }

        public void DDEPoke(int Channel, object Item, object Data)
        {
            Application.DDEPoke(Channel, Item, Data);
        }

        public dynamic DDERequest(int Channel, string Item)
        {
            return this.Application.DDERequest(Channel, Item);
        }

        public void DDETerminate(int Channel)
        {
            Application.DDETerminate(Channel);
        }

        public void DeleteChartAutoFormat(string Name)
        {
            Application.DeleteChartAutoFormat(Name);
        }

        public void DeleteCustomList(int ListNum)
        {
            Application.DeleteCustomList(ListNum);
        }

        public void DisplayXMLSourcePane(object XmlMap)
        {
            Application.DisplayXMLSourcePane(XmlMap);
        }

        public void DoubleClick()
        {
            Application.DoubleClick();
        }

        public dynamic Dummy1(object Arg1, object Arg2, object Arg3, object Arg4)
        {
            return this.Application.Dummy1(Arg1, Arg2, Arg3, Arg4);
        }

        public bool Dummy10(object arg)
        {
            return Application.Dummy10(arg);
        }

        public void Dummy11()
        {
            Application.Dummy11();
        }

        public void Dummy12(Excel.PivotTable p1, Excel.PivotTable p2)
        {
            Application.Dummy12(p1, p2);
        }

        public dynamic Dummy13(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return this.Application.Dummy13(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }

        public void Dummy14()
        {
            Application.Dummy14();
        }

        public dynamic Dummy2(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8)
        {
            return this.Application.Dummy2(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8);
        }

        public dynamic Dummy20(int grfCompareFunctions)
        {
            return this.Application.Dummy20(grfCompareFunctions);
        }

        public dynamic Dummy3()
        {
            return this.Application.Dummy3();
        }

        public dynamic Dummy4(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15)
        {
            return this.Application.Dummy4(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15);
        }

        public dynamic Dummy5(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13)
        {
            return this.Application.Dummy5(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13);
        }

        public dynamic Dummy6()
        {
            return this.Application.Dummy6();
        }

        public dynamic Dummy7()
        {
            return this.Application.Dummy7();
        }

        public dynamic Dummy8(object Arg1)
        {
            return this.Application.Dummy8(Arg1);
        }

        public dynamic Dummy9()
        {
            return this.Application.Dummy9();
        }

        public dynamic Evaluate(object Name)
        {
            return this.Application.Evaluate(Name);
        }

        public dynamic ExecuteExcel4Macro(string String)
        {
            return this.Application.ExecuteExcel4Macro(String);
        }

        public bool FindFile()
        {
            return Application.FindFile();
        }

        public dynamic GetCustomListContents(int ListNum)
        {
            return this.Application.GetCustomListContents(ListNum);
        }

        public int GetCustomListNum(object ListArray)
        {
            return Application.GetCustomListNum(ListArray);
        }

        public dynamic GetOpenFilename(object FileFilter, object FilterIndex, object Title, object ButtonText, object MultiSelect)
        {
            return this.Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect);
        }

        public string GetPhonetic(object Text)
        {
            return Application.GetPhonetic(Text);
        }

        public dynamic GetSaveAsFilename(object InitialFilename, object FileFilter, object FilterIndex, object Title, object ButtonText)
        {
            return this.Application.GetSaveAsFilename(InitialFilename, FileFilter, FilterIndex, Title, ButtonText);
        }

        public dynamic get_Caller(object Index)
        {
            return Application.Caller[Index];
        }

        public dynamic get_ClipboardFormats(object Index)
        {
            return Application.ClipboardFormats[Index];
        }

        public dynamic get_FileConverters(object Index1, object Index2)
        {
            return Application.FileConverters[Index1, Index2];
        }

        public FileDialog get_FileDialog(MsoFileDialogType fileDialogType)
        {
            return Application.FileDialog[fileDialogType];
        }

        public dynamic get_International(object Index)
        {
            return Application.International[Index];
        }

        public dynamic get_PreviousSelections(object Index)
        {
            return Application.PreviousSelections[Index];
        }

        public Excel.Range get_Range(object Cell1, object Cell2)
        {
            return Application.Range[Cell1, Cell2];
        }

        public dynamic get_RegisteredFunctions(object Index1, object Index2)
        {
            return Application.RegisteredFunctions[Index1, Index2];
        }

        public Excel.Menu get_ShortcutMenus(int Index)
        {
            return Application.ShortcutMenus[Index];
        }

        public void Goto(object Reference, object Scroll)
        {
            Application.Goto(Reference, Scroll);
        }

        public void Help(object HelpFile, object HelpContextID)
        {
            Application.Help(HelpFile, HelpContextID);
        }

        public double InchesToPoints(double Inches)
        {
            return Application.InchesToPoints(Inches);
        }

        public dynamic InputBox(string Prompt, object Title, object Default, object Left, object Top, object HelpFile, object HelpContextID, object Type)
        {
            return this.Application.InputBox(Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type);
        }

        public Excel.Range Intersect(Excel.Range Arg1, Excel.Range Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return Application.Intersect(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }

        public void MacroOptions(object Macro, object Description, object HasMenu, object MenuText, object HasShortcutKey, object ShortcutKey, object Category, object StatusBar, object HelpContextID, object HelpFile)
        {
            Application.MacroOptions(Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile);
        }

        public void MacroOptions2(object Macro, object Description, object HasMenu, object MenuText, object HasShortcutKey, object ShortcutKey, object Category, object StatusBar, object HelpContextID, object HelpFile, object ArgumentDescriptions)
        {
            Application.MacroOptions2(Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile, ArgumentDescriptions);
        }

        public void MailLogoff()
        {
            Application.MailLogoff();
        }

        public void MailLogon(object Name, object Password, object DownloadNewMail)
        {
            Application.MailLogon(Name, Password, DownloadNewMail);
        }

        public Excel.Workbook NextLetter()
        {
            return Application.NextLetter();
        }

        public void OnKey(string Key, object Procedure)
        {
            Application.OnKey(Key, Procedure);
        }

        public void OnRepeat(string Text, string Procedure)
        {
            Application.OnRepeat(Text, Procedure);
        }

        public void OnTime(object EarliestTime, string Procedure, object LatestTime, object Schedule)
        {
            Application.OnTime(EarliestTime, Procedure, LatestTime, Schedule);
        }

        public void OnUndo(string Text, string Procedure)
        {
            Application.OnUndo(Text, Procedure);
        }

        public void Quit()
        {
            Application.Quit();
        }

        public void RecordMacro(object BasicCode, object XlmCode)
        {
            Application.RecordMacro(BasicCode, XlmCode);
        }

        public bool RegisterXLL(string Filename)
        {
            return Application.RegisterXLL(Filename);
        }

        public void Repeat()
        {
            Application.Repeat();
        }

        public void ResetTipWizard()
        {
            Application.ResetTipWizard();
        }

        public dynamic Run(object Macro, object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return this.Application.Run(Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }

        public void Save(object Filename)
        {
            Application.Save(Filename);
        }

        public void SaveWorkspace(object Filename)
        {
            Application.SaveWorkspace(Filename);
        }

        public void SendKeys(object Keys, object Wait)
        {
            Application.SendKeys(Keys, Wait);
        }

        public void SetDefaultChart(object FormatName, object Gallery)
        {
            Application.SetDefaultChart(FormatName, Gallery);
        }

        public int SharePointVersion(string bstrUrl)
        {
            return Application.SharePointVersion(bstrUrl);
        }

        public dynamic Support(object Object, int ID, object arg)
        {
            return this.Application.Support(Object, ID, arg);
        }

        public void Undo()
        {
            Application.Undo();
        }

        public Excel.Range Union(Excel.Range Arg1, Excel.Range Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return Application.Union(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }

        public void Volatile(object Volatile)
        {
            Application.Volatile(Volatile);
        }

        public bool Wait(object Time)
        {
            return Application.Wait(Time);
        }

        public dynamic _Evaluate(object Name)
        {
            return this.Application._Evaluate(Name);
        }

        public void _FindFile()
        {
            Application._FindFile();
        }

        public dynamic _Run2(object Macro, object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return this.Application._Run2(Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }

        public void _Wait(object Time)
        {
            Application._Wait(Time);
        }

        public dynamic _WSFunction(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7, object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return this.Application._WSFunction(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
        }
        #endregion
    }
}
