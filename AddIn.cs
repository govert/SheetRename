using ExcelDna.Integration;

namespace SheetRename
{
    public class AddIn : IExcelAddIn
    {
        SheetNameMonitor _sheetNameMonitor;

        public void AutoOpen()
        {
            _sheetNameMonitor = new SheetNameMonitor();
            _sheetNameMonitor.SheetNameChanged += _sheetNameMonitor_SheetNameChanged;
            
        }

        private void _sheetNameMonitor_SheetNameChanged(object sender, SheetNameMonitor.SheetNameChangedArgs e)
        {
            string message = $"The worksheet name changed from {e.OldSheetName} to {e.NewSheetName}";
            System.Windows.Forms.MessageBox.Show(message);
        }

        public void AutoClose()
        {
        }
    }


}