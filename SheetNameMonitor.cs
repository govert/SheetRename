using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System;

namespace SheetRename
{
    internal class SheetNameMonitor
    {
        public class SheetNameChangedArgs
        {
            public string OldSheetName { get; set; }
            public string NewSheetName { get; set; }
        }

        readonly Application Application;

        string _activeSheetPreviousName;
        int _deactivatedSheetIndex;
        string _activeSheetName;

        public event EventHandler<SheetNameChangedArgs> SheetNameChanged;
        
        public SheetNameMonitor()
        {
            Application = ExcelDnaUtil.Application as Application;
            
            Application.SheetActivate += Application_SheetActivate;
            Application.SheetDeactivate += Application_SheetDeactivate;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.WorkbookDeactivate += Application_WorkbookDeactivate;
        }

        void Application_WorkbookOpen(Workbook Wb)
        {
            _activeSheetPreviousName = Wb.ActiveSheet.Name;
        }

        void Application_SheetActivate(object Sh)
        {
            SheetNameChange(1);
        }

        void Application_SheetDeactivate(object Sh)
        {
            _deactivatedSheetIndex = ((Worksheet)Sh).Index;
            SheetNameChange(0);
        }

        void Application_SheetSelectionChange(object Sh, Range Target)
        {
            SheetNameChange(2);
        }

        void Application_WorkbookActivate(Workbook Wb)
        {
            _activeSheetPreviousName = ((Worksheet)Wb.ActiveSheet).Name;
        }
        void Application_WorkbookDeactivate(Workbook Wb)
        {
            _deactivatedSheetIndex = ((Worksheet)Wb.ActiveSheet).Index;
            SheetNameChange(2);

            _activeSheetPreviousName = null;
            _deactivatedSheetIndex = 0;
            _activeSheetName = null;
        }

        void SheetNameChange(int caller)
        {
            Workbook activeWorkbook = Application.ActiveWorkbook;
            Worksheet activeSheet = activeWorkbook.ActiveSheet;

            switch (caller)
            {
                case 0:

                    if (_deactivatedSheetIndex != activeSheet.Index)
                    {
                        Worksheet deActivatedSheet = activeWorkbook.Worksheets[_deactivatedSheetIndex];

                        if (deActivatedSheet.Name != _activeSheetPreviousName)
                        {
                            OnSheetNameChanged(_activeSheetPreviousName, deActivatedSheet.Name);
                            _activeSheetPreviousName = activeSheet.Name;
                        }
                    }

                    break;
                case 1:
                    _activeSheetPreviousName = activeSheet.Name;
                    _activeSheetName = activeSheet.Name;
                    break;
                case 2:
                    _activeSheetName = activeSheet.Name;

                    if (_activeSheetName != _activeSheetPreviousName)
                    {
                        OnSheetNameChanged(_activeSheetPreviousName, _activeSheetName);
                        _activeSheetPreviousName = _activeSheetName;
                    }
                    break;
            }
        }

        protected virtual void OnSheetNameChanged(string oldName, string newName)
        {
            var e = new SheetNameChangedArgs { OldSheetName = oldName, NewSheetName = newName };
            var handler = SheetNameChanged;
            handler?.Invoke(this, e);
        }

    }
}
