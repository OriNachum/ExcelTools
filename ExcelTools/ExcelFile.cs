using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace ExcelTools
{
    public class ExcelFile : Excel.Workbook, IDisposable
    {
        public ExcelFile(Excel.Application pApp, string pFileAddress, bool pReadOnly=true, bool pEditable=false, bool pNew = false)
        {
            App = pApp;

            Readonly = pReadOnly;

            if (pNew)
            {
                Worksheet = pApp.Application.Workbooks.Add();
            }
            else
            {
                Worksheet = pApp.Workbooks.Open(pFileAddress, ReadOnly: pReadOnly, Editable: pEditable);
                if (ReadOnly)
                {
                    pApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                var Calculation = pApp.Calculation;
            }
        }

        public Excel.Workbook Worksheet { get; private set; }

        public void Dispose()
        {
            Worksheet.Close(SaveChanges: false);
            try
            {
                int result;
                do
                {
                    result = Worksheet.ReleaseObject();
                }
                while (result > 0);
            }
            catch (Exception ex)
            {
                Worksheet = null;
                throw new Exception("Error releasing object: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
            Worksheet = null;
        }

        #region Workbook
        public bool AcceptLabelsInFormulas
        {
            get
            {
                return Worksheet.AcceptLabelsInFormulas;
            }

            set
            {
                Worksheet.AcceptLabelsInFormulas = value;
            }
        }

        public int AccuracyVersion
        {
            get
            {
                return Worksheet.AccuracyVersion;
            }

            set
            {
                Worksheet.AccuracyVersion = value;
            }
        }

        public Excel.Chart ActiveChart
        {
            get
            {
                return Worksheet.ActiveChart;
            }
        }

        public dynamic ActiveSheet
        {
            get
            {
                return Worksheet.ActiveSheet;
            }
        }

        public Excel.Slicer ActiveSlicer
        {
            get
            {
                return Worksheet.ActiveSlicer;
            }
        }

        public Excel.Application Application
        {
            get
            {
                return Worksheet.Application;
            }
        }

        public string Author
        {
            get
            {
                return Worksheet.Author;
            }

            set
            {
                Worksheet.Author = value;
            }
        }

        public int AutoUpdateFrequency
        {
            get
            {
                return Worksheet.AutoUpdateFrequency;
            }

            set
            {
                Worksheet.AutoUpdateFrequency = value;
            }
        }

        public bool AutoUpdateSaveChanges
        {
            get
            {
                return Worksheet.AutoUpdateSaveChanges;
            }

            set
            {
                Worksheet.AutoUpdateSaveChanges = value;
            }
        }

        public dynamic BuiltinDocumentProperties
        {
            get
            {
                return Worksheet.BuiltinDocumentProperties;
            }
        }

        public int CalculationVersion
        {
            get
            {
                return Worksheet.CalculationVersion;
            }
        }

        public int ChangeHistoryDuration
        {
            get
            {
                return Worksheet.ChangeHistoryDuration;
            }

            set
            {
                Worksheet.ChangeHistoryDuration = value;
            }
        }

        public Excel.Sheets Charts
        {
            get
            {
                return Worksheet.Charts;
            }
        }

        public bool CheckCompatibility
        {
            get
            {
                return Worksheet.CheckCompatibility;
            }

            set
            {
                Worksheet.CheckCompatibility = value;
            }
        }

        public string CodeName
        {
            get
            {
                return Worksheet.CodeName;
            }
        }

        public CommandBars CommandBars
        {
            get
            {
                return Worksheet.CommandBars;
            }
        }

        public string Comments
        {
            get
            {
                return Worksheet.Comments;
            }

            set
            {
                Worksheet.Comments = value;
            }
        }

        public Excel.XlSaveConflictResolution ConflictResolution
        {
            get
            {
                return Worksheet.ConflictResolution;
            }

            set
            {
                Worksheet.ConflictResolution = value;
            }
        }

        public Excel.Connections Connections
        {
            get
            {
                return Worksheet.Connections;
            }
        }

        public bool ConnectionsDisabled
        {
            get
            {
                return Worksheet.ConnectionsDisabled;
            }
        }

        public dynamic Container
        {
            get
            {
                return Worksheet.Container;
            }
        }

        public MetaProperties ContentTypeProperties
        {
            get
            {
                return Worksheet.ContentTypeProperties;
            }
        }

        public bool CreateBackup
        {
            get
            {
                return Worksheet.CreateBackup;
            }
        }

        public Excel.XlCreator Creator
        {
            get
            {
                return Worksheet.Creator;
            }
        }

        public dynamic CustomDocumentProperties
        {
            get
            {
                return Worksheet.CustomDocumentProperties;
            }
        }

        public Excel.CustomViews CustomViews
        {
            get
            {
                return Worksheet.CustomViews;
            }
        }

        public CustomXMLParts CustomXMLParts
        {
            get
            {
                return Worksheet.CustomXMLParts;
            }
        }

        public bool Date1904
        {
            get
            {
                return Worksheet.Date1904;
            }

            set
            {
                Worksheet.Date1904 = value;
            }
        }

        public dynamic DefaultPivotTableStyle
        {
            get
            {
                return Worksheet.DefaultPivotTableStyle;
            }

            set
            {
                Worksheet.DefaultPivotTableStyle = value;
            }
        }

        public dynamic DefaultSlicerStyle
        {
            get
            {
                return Worksheet.DefaultSlicerStyle;
            }

            set
            {
                Worksheet.DefaultSlicerStyle = value;
            }
        }

        public dynamic DefaultTableStyle
        {
            get
            {
                return Worksheet.DefaultTableStyle;
            }

            set
            {
                Worksheet.DefaultTableStyle = value;
            }
        }

        public Excel.Sheets DialogSheets
        {
            get
            {
                return Worksheet.DialogSheets;
            }
        }

        public Excel.XlDisplayDrawingObjects DisplayDrawingObjects
        {
            get
            {
                return Worksheet.DisplayDrawingObjects;
            }

            set
            {
                Worksheet.DisplayDrawingObjects = value;
            }
        }

        public bool DisplayInkComments
        {
            get
            {
                return Worksheet.DisplayInkComments;
            }

            set
            {
                Worksheet.DisplayInkComments = value;
            }
        }

        public DocumentInspectors DocumentInspectors
        {
            get
            {
                return Worksheet.DocumentInspectors;
            }
        }

        public DocumentLibraryVersions DocumentLibraryVersions
        {
            get
            {
                return Worksheet.DocumentLibraryVersions;
            }
        }

        public bool DoNotPromptForConvert
        {
            get
            {
                return Worksheet.DoNotPromptForConvert;
            }

            set
            {
                Worksheet.DoNotPromptForConvert = value;
            }
        }

        public bool EnableAutoRecover
        {
            get
            {
                return Worksheet.EnableAutoRecover;
            }

            set
            {
                Worksheet.EnableAutoRecover = value;
            }
        }

        public string EncryptionProvider
        {
            get
            {
                return Worksheet.EncryptionProvider;
            }

            set
            {
                Worksheet.EncryptionProvider = value;
            }
        }

        public bool EnvelopeVisible
        {
            get
            {
                return Worksheet.EnvelopeVisible;
            }

            set
            {
                Worksheet.EnvelopeVisible = value;
            }
        }

        public Excel.Sheets Excel4IntlMacroSheets
        {
            get
            {
                return Worksheet.Excel4IntlMacroSheets;
            }
        }

        public Excel.Sheets Excel4MacroSheets
        {
            get
            {
                return Worksheet.Excel4MacroSheets;
            }
        }

        public bool Excel8CompatibilityMode
        {
            get
            {
                return Worksheet.Excel8CompatibilityMode;
            }
        }

        public Excel.XlFileFormat FileFormat
        {
            get
            {
                return Worksheet.FileFormat;
            }
        }

        public bool Final
        {
            get
            {
                return Worksheet.Final;
            }

            set
            {
                Worksheet.Final = value;
            }
        }

        public bool ForceFullCalculation
        {
            get
            {
                return Worksheet.ForceFullCalculation;
            }

            set
            {
                Worksheet.ForceFullCalculation = value;
            }
        }

        public string FullName
        {
            get
            {
                return Worksheet.FullName;
            }
        }

        public string FullNameURLEncoded
        {
            get
            {
                return Worksheet.FullNameURLEncoded;
            }
        }

        public bool HasMailer
        {
            get
            {
                return Worksheet.HasMailer;
            }

            set
            {
                Worksheet.HasMailer = value;
            }
        }

        public bool HasPassword
        {
            get
            {
                return Worksheet.HasPassword;
            }
        }

        public bool HasRoutingSlip
        {
            get
            {
                return Worksheet.HasRoutingSlip;
            }

            set
            {
                Worksheet.HasRoutingSlip = value;
            }
        }

        public bool HasVBProject
        {
            get
            {
                return Worksheet.HasVBProject;
            }
        }

        public bool HighlightChangesOnScreen
        {
            get
            {
                return Worksheet.HighlightChangesOnScreen;
            }

            set
            {
                Worksheet.HighlightChangesOnScreen = value;
            }
        }

        public HTMLProject HTMLProject
        {
            get
            {
                return Worksheet.HTMLProject;
            }
        }

        public Excel.IconSets IconSets
        {
            get
            {
                return Worksheet.IconSets;
            }
        }

        public bool InactiveListBorderVisible
        {
            get
            {
                return Worksheet.InactiveListBorderVisible;
            }

            set
            {
                Worksheet.InactiveListBorderVisible = value;
            }
        }

        public bool IsAddin
        {
            get
            {
                return Worksheet.IsAddin;
            }

            set
            {
                Worksheet.IsAddin = value;
            }
        }

        public bool IsInplace
        {
            get
            {
                return Worksheet.IsInplace;
            }
        }

        public bool KeepChangeHistory
        {
            get
            {
                return Worksheet.KeepChangeHistory;
            }

            set
            {
                Worksheet.KeepChangeHistory = value;
            }
        }

        public string Keywords
        {
            get
            {
                return Worksheet.Keywords;
            }

            set
            {
                Worksheet.Keywords = value;
            }
        }

        public bool ListChangesOnNewSheet
        {
            get
            {
                return Worksheet.ListChangesOnNewSheet;
            }

            set
            {
                Worksheet.ListChangesOnNewSheet = value;
            }
        }

        public Excel.Mailer Mailer
        {
            get
            {
                return Worksheet.Mailer;
            }
        }

        public Excel.Sheets Modules
        {
            get
            {
                return Worksheet.Modules;
            }
        }

        public bool MultiUserEditing
        {
            get
            {
                return Worksheet.MultiUserEditing;
            }
        }

        public string Name
        {
            get
            {
                return Worksheet.Name;
            }
        }

        public Excel.Names Names
        {
            get
            {
                return Worksheet.Names;
            }
        }

        public string OnSave
        {
            get
            {
                return Worksheet.OnSave;
            }

            set
            {
                Worksheet.OnSave = value;
            }
        }

        public string OnSheetActivate
        {
            get
            {
                return Worksheet.OnSheetActivate;
            }

            set
            {
                Worksheet.OnSheetActivate = value;
            }
        }

        public string OnSheetDeactivate
        {
            get
            {
                return Worksheet.OnSheetDeactivate;
            }

            set
            {
                Worksheet.OnSheetDeactivate = value;
            }
        }

        public dynamic Parent
        {
            get
            {
                return Worksheet.Parent;
            }
        }

        public string Password
        {
            get
            {
                return Worksheet.Password;
            }

            set
            {
                Worksheet.Password = value;
            }
        }

        public string PasswordEncryptionAlgorithm
        {
            get
            {
                return Worksheet.PasswordEncryptionAlgorithm;
            }
        }

        public bool PasswordEncryptionFileProperties
        {
            get
            {
                return Worksheet.PasswordEncryptionFileProperties;
            }
        }

        public int PasswordEncryptionKeyLength
        {
            get
            {
                return Worksheet.PasswordEncryptionKeyLength;
            }
        }

        public string PasswordEncryptionProvider
        {
            get
            {
                return Worksheet.PasswordEncryptionProvider;
            }
        }

        public string Path
        {
            get
            {
                return Worksheet.Path;
            }
        }

        public Permission Permission
        {
            get
            {
                return Worksheet.Permission;
            }
        }

        public bool PersonalViewListSettings
        {
            get
            {
                return Worksheet.PersonalViewListSettings;
            }

            set
            {
                Worksheet.PersonalViewListSettings = value;
            }
        }

        public bool PersonalViewPrintSettings
        {
            get
            {
                return Worksheet.PersonalViewPrintSettings;
            }

            set
            {
                Worksheet.PersonalViewPrintSettings = value;
            }
        }

        public bool PrecisionAsDisplayed
        {
            get
            {
                return Worksheet.PrecisionAsDisplayed;
            }

            set
            {
                Worksheet.PrecisionAsDisplayed = value;
            }
        }

        public bool ProtectStructure
        {
            get
            {
                return Worksheet.ProtectStructure;
            }
        }

        public bool ProtectWindows
        {
            get
            {
                return Worksheet.ProtectWindows;
            }
        }

        public Excel.PublishObjects PublishObjects
        {
            get
            {
                return Worksheet.PublishObjects;
            }
        }

        public bool ReadOnly
        {
            get
            {
                return Worksheet.ReadOnly;
            }
        }

        public bool ReadOnlyRecommended
        {
            get
            {
                return Worksheet.ReadOnlyRecommended;
            }

            set
            {
                Worksheet.ReadOnlyRecommended = value;
            }
        }

        public bool RemovePersonalInformation
        {
            get
            {
                return Worksheet.RemovePersonalInformation;
            }

            set
            {
                Worksheet.RemovePersonalInformation = value;
            }
        }

        public Excel.Research Research
        {
            get
            {
                return Worksheet.Research;
            }
        }

        public int RevisionNumber
        {
            get
            {
                return Worksheet.RevisionNumber;
            }
        }

        public bool Routed
        {
            get
            {
                return Worksheet.Routed;
            }
        }

        public Excel.RoutingSlip RoutingSlip
        {
            get
            {
                return Worksheet.RoutingSlip;
            }
        }

        public bool Saved
        {
            get
            {
                return Worksheet.Saved;
            }

            set
            {
                Worksheet.Saved = value;
            }
        }

        public bool SaveLinkValues
        {
            get
            {
                return Worksheet.SaveLinkValues;
            }

            set
            {
                Worksheet.SaveLinkValues = value;
            }
        }

        public ServerPolicy ServerPolicy
        {
            get
            {
                return Worksheet.ServerPolicy;
            }
        }

        public Excel.ServerViewableItems ServerViewableItems
        {
            get
            {
                return Worksheet.ServerViewableItems;
            }
        }

        public SharedWorkspace SharedWorkspace
        {
            get
            {
                return Worksheet.SharedWorkspace;
            }
        }

        public Excel.Sheets Sheets
        {
            get
            {
                return Worksheet.Sheets;
            }
        }

        public bool ShowConflictHistory
        {
            get
            {
                return Worksheet.ShowConflictHistory;
            }

            set
            {
                Worksheet.ShowConflictHistory = value;
            }
        }

        public bool ShowPivotChartActiveFields
        {
            get
            {
                return Worksheet.ShowPivotChartActiveFields;
            }

            set
            {
                Worksheet.ShowPivotChartActiveFields = value;
            }
        }

        public bool ShowPivotTableFieldList
        {
            get
            {
                return Worksheet.ShowPivotTableFieldList;
            }

            set
            {
                Worksheet.ShowPivotTableFieldList = value;
            }
        }

        public SignatureSet Signatures
        {
            get
            {
                return Worksheet.Signatures;
            }
        }

        public Excel.SlicerCaches SlicerCaches
        {
            get
            {
                return Worksheet.SlicerCaches;
            }
        }

        public SmartDocument SmartDocument
        {
            get
            {
                return Worksheet.SmartDocument;
            }
        }

        public Excel.SmartTagOptions SmartTagOptions
        {
            get
            {
                return Worksheet.SmartTagOptions;
            }
        }

        public Excel.Styles Styles
        {
            get
            {
                return Worksheet.Styles;
            }
        }

        public string Subject
        {
            get
            {
                return Worksheet.Subject;
            }

            set
            {
                Worksheet.Subject = value;
            }
        }

        public Sync Sync
        {
            get
            {
                throw new NotImplementedException();
                //return Worksheet.Sync;
            }
        }

        public Excel.TableStyles TableStyles
        {
            get
            {
                return Worksheet.TableStyles;
            }
        }

        public bool TemplateRemoveExtData
        {
            get
            {
                return Worksheet.TemplateRemoveExtData;
            }

            set
            {
                Worksheet.TemplateRemoveExtData = value;
            }
        }

        public OfficeTheme Theme
        {
            get
            {
                return Worksheet.Theme;
            }
        }

        public string Title
        {
            get
            {
                return Worksheet.Title;
            }

            set
            {
                Worksheet.Title = value;
            }
        }

        public Excel.XlUpdateLinks UpdateLinks
        {
            get
            {
                return Worksheet.UpdateLinks;
            }

            set
            {
                Worksheet.UpdateLinks = value;
            }
        }

        public bool UpdateRemoteReferences
        {
            get
            {
                return Worksheet.UpdateRemoteReferences;
            }

            set
            {
                Worksheet.UpdateRemoteReferences = value;
            }
        }

        public bool UserControl
        {
            get
            {
                return Worksheet.UserControl;
            }

            set
            {
                Worksheet.UserControl = value;
            }
        }

        public dynamic UserStatus
        {
            get
            {
                return Worksheet.UserStatus;
            }
        }

        public bool VBASigned
        {
            get
            {
                return Worksheet.VBASigned;
            }
        }

        public VBProject VBProject
        {
            get
            {
                return Worksheet.VBProject;
            }
        }

        public Excel.WebOptions WebOptions
        {
            get
            {
                return Worksheet.WebOptions;
            }
        }

        public Excel.Windows Windows
        {
            get
            {
                return Worksheet.Windows;
            }
        }


        public Excel.Sheets Worksheets
        {
            get
            {
                return Worksheet.Worksheets;
            }
        }

        public string WritePassword
        {
            get
            {
                return Worksheet.WritePassword;
            }

            set
            {
                Worksheet.WritePassword = value;
            }
        }

        public bool WriteReserved
        {
            get
            {
                return Worksheet.WriteReserved;
            }
        }

        public string WriteReservedBy
        {
            get
            {
                return Worksheet.WriteReservedBy;
            }
        }

        public Excel.XmlMaps XmlMaps
        {
            get
            {
                return Worksheet.XmlMaps;
            }
        }

        public Excel.XmlNamespaces XmlNamespaces
        {
            get
            {
                return Worksheet.XmlNamespaces;
            }
        }

        public string _CodeName
        {
            get
            {
                return Worksheet._CodeName;
            }

            set
            {
                Worksheet._CodeName = value;
            }
        }

        public bool _ReadOnlyRecommended
        {
            get
            {
                return Worksheet._ReadOnlyRecommended;
            }
        }

        public event Excel.WorkbookEvents_AddinInstallEventHandler AddinInstall;
        public event Excel.WorkbookEvents_AddinUninstallEventHandler AddinUninstall;
        public event Excel.WorkbookEvents_AfterSaveEventHandler AfterSave;
        public event Excel.WorkbookEvents_AfterXmlExportEventHandler AfterXmlExport;
        public event Excel.WorkbookEvents_AfterXmlImportEventHandler AfterXmlImport;
        public event Excel.WorkbookEvents_BeforeCloseEventHandler BeforeClose;
        public event Excel.WorkbookEvents_BeforePrintEventHandler BeforePrint;
        public event Excel.WorkbookEvents_BeforeSaveEventHandler BeforeSave;
        public event Excel.WorkbookEvents_BeforeXmlExportEventHandler BeforeXmlExport;
        public event Excel.WorkbookEvents_BeforeXmlImportEventHandler BeforeXmlImport;
        public event Excel.WorkbookEvents_DeactivateEventHandler Deactivate;
        public event Excel.WorkbookEvents_NewChartEventHandler NewChart;
        public event Excel.WorkbookEvents_NewSheetEventHandler NewSheet;
        public event Excel.WorkbookEvents_OpenEventHandler Open;
        public event Excel.WorkbookEvents_PivotTableCloseConnectionEventHandler PivotTableCloseConnection;
        public event Excel.WorkbookEvents_PivotTableOpenConnectionEventHandler PivotTableOpenConnection;
        public event Excel.WorkbookEvents_RowsetCompleteEventHandler RowsetComplete;
        public event Excel.WorkbookEvents_SheetActivateEventHandler SheetActivate;
        public event Excel.WorkbookEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        public event Excel.WorkbookEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        public event Excel.WorkbookEvents_SheetCalculateEventHandler SheetCalculate;
        public event Excel.WorkbookEvents_SheetChangeEventHandler SheetChange;
        public event Excel.WorkbookEvents_SheetDeactivateEventHandler SheetDeactivate;
        public event Excel.WorkbookEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        public event Excel.WorkbookEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChange;
        public event Excel.WorkbookEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChanges;
        public event Excel.WorkbookEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChanges;
        public event Excel.WorkbookEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChanges;
        public event Excel.WorkbookEvents_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSync;
        public event Excel.WorkbookEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        public event Excel.WorkbookEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        public event Excel.WorkbookEvents_WindowActivateEventHandler WindowActivate;
        public event Excel.WorkbookEvents_WindowDeactivateEventHandler WindowDeactivate;
        public event Excel.WorkbookEvents_WindowResizeEventHandler WindowResize;

        event Excel.WorkbookEvents_ActivateEventHandler Excel.WorkbookEvents_Event.Activate
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

        event Excel.WorkbookEvents_SyncEventHandler Excel.WorkbookEvents_Event.Sync
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

        public void AcceptAllChanges(object When, object Who, object Where)
        {
            Worksheet.AcceptAllChanges(When, Who, Where);
        }

        public void Activate()
        {
            Worksheet.Activate();
        }

        public void AddToFavorites()
        {
            Worksheet.AddToFavorites();
        }

        public void ApplyTheme(string Filename)
        {
            Worksheet.ApplyTheme(Filename);
        }

        public void BreakLink(string Name, Excel.XlLinkType Type)
        {
            Worksheet.BreakLink(Name, Type);
        }

        public bool CanCheckIn()
        {
            return Worksheet.CanCheckIn();
        }

        public void ChangeFileAccess(Excel.XlFileAccess Mode, object WritePassword, object Notify)
        {
            Worksheet.ChangeFileAccess(Mode, WritePassword, Notify);
        }

        public void ChangeLink(string Name, string NewName, Excel.XlLinkType Type = Excel.XlLinkType.xlLinkTypeExcelLinks)
        {
            Worksheet.ChangeLink(Name, NewName, Type);
        }

        public void CheckIn(object SaveChanges, object Comments, object MakePublic)
        {
            Worksheet.CheckIn(SaveChanges, Comments, MakePublic);
        }

        public void CheckInWithVersion(object SaveChanges, object Comments, object MakePublic, object VersionType)
        {
            Worksheet.CheckInWithVersion(SaveChanges, Comments, MakePublic, VersionType);
        }

        public void Close(object SaveChanges, object Filename, object RouteWorkbook)
        {
            Worksheet.Close(SaveChanges, Filename, RouteWorkbook);
        }

        public void DeleteNumberFormat(string NumberFormat)
        {
            Worksheet.DeleteNumberFormat(NumberFormat);
        }

        public void Dummy16()
        {
            Worksheet.Dummy16();
        }

        public void Dummy17(int calcid)
        {
            Worksheet.Dummy17(calcid);
        }

        public void Dummy26()
        {
            Worksheet.Dummy26();
        }

        public void Dummy27()
        {
            Worksheet.Dummy27();
        }

        public void EnableConnections()
        {
            Worksheet.EnableConnections();
        }

        public void EndReview()
        {
            Worksheet.EndReview();
        }

        public bool ExclusiveAccess()
        {
            return Worksheet.ExclusiveAccess();
        }

        public void ExportAsFixedFormat(Excel.XlFixedFormatType Type, object Filename, object Quality, object IncludeDocProperties, object IgnorePrintAreas, object From, object To, object OpenAfterPublish, object FixedFormatExtClassPtr)
        {
            Worksheet.ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr);
        }

        public void FollowHyperlink(string Address, object SubAddress, object NewWindow, object AddHistory, object ExtraInfo, object Method, object HeaderInfo)
        {
            Worksheet.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo);
        }

        public void ForwardMailer()
        {
            Worksheet.ForwardMailer();
        }

        public WorkflowTasks GetWorkflowTasks()
        {
            return Worksheet.GetWorkflowTasks();
        }

        public WorkflowTemplates GetWorkflowTemplates()
        {
            return Worksheet.GetWorkflowTemplates();
        }

        public dynamic get_Colors(object Index)
        {
            return Worksheet.Colors[Index];
        }

        public void HighlightChangesOptions(object When, object Who, object Where)
        {
            Worksheet.HighlightChangesOptions(When, Who, Where);
        }

        public dynamic LinkInfo(string Name, Excel.XlLinkInfo LinkInfo, object Type, object EditionRef)
        {
            return this.Worksheet.LinkInfo(Name, LinkInfo, Type, EditionRef);
        }

        public dynamic LinkSources(object Type)
        {
            return this.Worksheet.LinkSources(Type);
        }

        public void LockServerFile()
        {
            Worksheet.LockServerFile();
        }

        public void MergeWorkbook(object Filename)
        {
            Worksheet.MergeWorkbook(Filename);
        }

        public Excel.Window NewWindow()
        {
            return Worksheet.NewWindow();
        }

        public void OpenLinks(string Name, object ReadOnly, object Type)
        {
            Worksheet.OpenLinks(Name, ReadOnly, Type);
        }

        public Excel.PivotCaches PivotCaches()
        {
            return Worksheet.PivotCaches();
        }

        public void PivotTableWizard(object SourceType, object SourceData, object TableDestination, object TableName, object RowGrand, object ColumnGrand, object SaveData, object HasAutoFormat, object AutoPage, object Reserved, object BackgroundQuery, object OptimizeCache, object PageFieldOrder, object PageFieldWrapCount, object ReadData, object Connection)
        {
            Worksheet.PivotTableWizard(SourceType, SourceData, TableDestination, TableName, RowGrand, ColumnGrand, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection);
        }

        public void Post(object DestName)
        {
            Worksheet.Post(DestName);
        }

        public void PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            Worksheet.PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName);
        }

        public void PrintOutEx(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName, object IgnorePrintAreas)
        {
            Worksheet.PrintOutEx(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas);
        }

        public void PrintPreview(object EnableChanges)
        {
            Worksheet.PrintPreview(EnableChanges);
        }

        public void Protect(object Password, object Structure, object Windows)
        {
            Worksheet.Protect(Password, Structure, Windows);
        }

        public void ProtectSharing(object Filename, object Password, object WriteResPassword, object ReadOnlyRecommended, object CreateBackup, object SharingPassword)
        {
            Worksheet.ProtectSharing(Filename, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword);
        }

        public void ProtectSharingEx(object Filename, object Password, object WriteResPassword, object ReadOnlyRecommended, object CreateBackup, object SharingPassword, object FileFormat)
        {
            Worksheet.ProtectSharingEx(Filename, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword, FileFormat);
        }

        public void PurgeChangeHistoryNow(int Days, object SharingPassword)
        {
            Worksheet.PurgeChangeHistoryNow(Days, SharingPassword);
        }

        public void RecheckSmartTags()
        {
            Worksheet.RecheckSmartTags();
        }

        public void RefreshAll()
        {
            Worksheet.RefreshAll();
        }

        public void RejectAllChanges(object When, object Who, object Where)
        {
            Worksheet.RejectAllChanges(When, Who, Where);
        }

        public void ReloadAs(MsoEncoding Encoding)
        {
            Worksheet.ReloadAs(Encoding);
        }

        public void RemoveDocumentInformation(Excel.XlRemoveDocInfoType RemoveDocInfoType)
        {
            Worksheet.RemoveDocumentInformation(RemoveDocInfoType);
        }

        public void RemoveUser(int Index)
        {
            Worksheet.RemoveUser(Index);
        }

        public void Reply()
        {
            Worksheet.Reply();
        }

        public void ReplyAll()
        {
            Worksheet.ReplyAll();
        }

        public void ReplyWithChanges(object ShowMessage)
        {
            Worksheet.ReplyWithChanges(ShowMessage);
        }

        public void ResetColors()
        {
            Worksheet.ResetColors();
        }

        public void Route()
        {
            Worksheet.Route();
        }

        public void RunAutoMacros(Excel.XlRunAutoMacro Which)
        {
            Worksheet.RunAutoMacros(Which);
        }

        public void Save()
        {
            Worksheet.Save();
        }

        public void SaveAs(object Filename, object FileFormat, object Password, object WriteResPassword, object ReadOnlyRecommended, object CreateBackup, Excel.XlSaveAsAccessMode AccessMode = Excel.XlSaveAsAccessMode.xlNoChange, object ConflictResolution = null, object AddToMru = null, object TextCodepage = null, object TextVisualLayout = null, object Local = null)
        {
            Worksheet.SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local);
        }

        public void SaveAsXMLData(string Filename, Excel.XmlMap Map)
        {
            Worksheet.SaveAsXMLData(Filename, Map);
        }

        public void SaveCopyAs(object Filename)
        {
            Worksheet.SaveCopyAs(Filename);
        }

        public void sblt(string s)
        {
            Worksheet.sblt(s);
        }

        public void SendFaxOverInternet(object Recipients, object Subject, object ShowMessage)
        {
            Worksheet.SendFaxOverInternet(Recipients, Subject, ShowMessage);
        }

        public void SendForReview(object Recipients, object Subject, object ShowMessage, object IncludeAttachment)
        {
            Worksheet.SendForReview(Recipients, Subject, ShowMessage, IncludeAttachment);
        }

        public void SendMail(object Recipients, object Subject, object ReturnReceipt)
        {
            Worksheet.SendMail(Recipients, Subject, ReturnReceipt);
        }

        public void SendMailer(object FileFormat, Excel.XlPriority Priority = Excel.XlPriority.xlPriorityNormal)
        {
            Worksheet.SendMailer(FileFormat, Priority);
        }

        public void SetLinkOnData(string Name, object Procedure)
        {
            Worksheet.SetLinkOnData(Name, Procedure);
        }

        public void SetPasswordEncryptionOptions(object PasswordEncryptionProvider, object PasswordEncryptionAlgorithm, object PasswordEncryptionKeyLength, object PasswordEncryptionFileProperties)
        {
            Worksheet.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties);
        }

        public void set_Colors(object Index, object RHS)
        {
            Worksheet.Colors[Index] = RHS;
        }

        public void ToggleFormsDesign()
        {
            Worksheet.ToggleFormsDesign();
        }

        public void Unprotect(object Password)
        {
            Worksheet.Unprotect(Password);
        }

        public void UnprotectSharing(object SharingPassword)
        {
            Worksheet.UnprotectSharing(SharingPassword);
        }

        public void UpdateFromFile()
        {
            Worksheet.UpdateFromFile();
        }

        public void UpdateLink(object Name, object Type)
        {
            Worksheet.UpdateLink(Name, Type);
        }

        public void WebPagePreview()
        {
            Worksheet.WebPagePreview();
        }

        public Excel.XlXmlImportResult XmlImport(string Url, out Excel.XmlMap ImportMap, object Overwrite, object Destination)
        {
            return Worksheet.XmlImport(Url, out ImportMap, Overwrite, Destination);
        }

        public Excel.XlXmlImportResult XmlImportXml(string Data, out Excel.XmlMap ImportMap, object Overwrite, object Destination)
        {
            return Worksheet.XmlImportXml(Data, out ImportMap, Overwrite, Destination);
        }

        public void _PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate)
        {
            Worksheet._PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate);
        }

        public void _Protect(object Password, object Structure, object Windows)
        {
            Worksheet._Protect(Password, Structure, Windows);
        }

        public void _SaveAs(object Filename, object FileFormat, object Password, object WriteResPassword, object ReadOnlyRecommended, object CreateBackup, Excel.XlSaveAsAccessMode AccessMode = Excel.XlSaveAsAccessMode.xlNoChange, object ConflictResolution = null, object AddToMru = null, object TextCodepage = null, object TextVisualLayout = null)
        {
            Worksheet._SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout);
        }
        #endregion

        readonly Excel.Application App;
        readonly bool Readonly;
        readonly Excel.XlCalculation Calculation; 
    }
}
