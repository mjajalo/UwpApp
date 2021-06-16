using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Runtime.CompilerServices;
using myapp_uwp.Models;
using Windows.Storage;
using System.IO;
using Windows.Storage.Pickers;
using Windows.UI.Popups;
using pdftron.Filters;
using pdftron.PDF;
using System.Windows.Input;
using Windows.ApplicationModel;
using Windows.ApplicationModel.DataTransfer;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Input;

namespace myapp_uwp.ViewModels
{
    class MainViewModel : BaseViewModel
    {
        private MainModel _mainModel = new MainModel();

        public MainViewModel()
        {
            CMDOpenFile = new RelayCommand(OpenFile);
            CMDConvertFile = new RelayCommand(ConvertFile);
            CMDGoNextPage = new RelayCommand(GoNextPage);
            CMDGoPrevPage = new RelayCommand(GoPrevPage);
            CMDZoomIn = new RelayCommand(ZoomIn);
            CMDZoomOut = new RelayCommand(ZoomOut);
            CMDResetZoom = new RelayCommand(ResetZoom);
            CMDCloseFile = new RelayCommand(CloseFile);
            pdftron.PDFNet.AddResourceSearchPath(System.IO.Path.Combine(Package.Current.InstalledLocation.Path, "Resources"));          
            _myToolManager = new pdftron.PDF.Tools.ToolManager(MyPDFViewCtrl);
            MyPDFViewCtrl.OnPageNumberChanged += MyPDFViewCtrl_OnPageNumberChanged;
        }

        private void MyPDFViewCtrl_OnPageNumberChanged(int current_page, int num_pages)
        {
            NotifyPropertyChanged("CurrentPage");
        }

        private void MyPDFViewCtrl_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = Windows.ApplicationModel.DataTransfer.DataPackageOperation.Copy;
            e.DragUIOverride.Caption = "Open PDF file";
        }

        private async void MyPDFViewCtrl_DropAsync(object sender, DragEventArgs e)
        {
            if (e.DataView.Contains(StandardDataFormats.StorageItems))
            {
                IReadOnlyList<IStorageItem> items = await e.DataView.GetStorageItemsAsync();
                if (items.Count == 1)
                {
                    StorageFile storageFile = items[0] as StorageFile;

                    if (storageFile.FileType.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        // Note: Drag and drop on UWP only allows Read-Only access
                        await DragOpenFile(storageFile);
                    }
                }
            }
        }

        public string FileName
        {
            get { return _mainModel.FileName; }
            set
            {
                if (_mainModel.FileName != value)
                {
                    _mainModel.FileName = value;
                    NotifyPropertyChanged("FileName");
                    NotifyPropertyChanged("IsDocOpen");
                }
            }
        }

        public int PageCount
        {
            get { return _mainModel.PageCount; }
            set
            {
                if (_mainModel.PageCount != value)
                {
                    _mainModel.PageCount = value;
                    NotifyPropertyChanged("PageCount");
                    NotifyPropertyChanged("CurrentPage");
                }
            }
        }

        public int CurrentPage
        {
            get { return MyPDFViewCtrl.GetCurrentPage(); }
        }

        public double ZoomAmount
        {
            get { return _mainModel.ZoomAmount; }
            set
            {
                if (_mainModel.ZoomAmount != value)
                {
                    _mainModel.ZoomAmount = value;                  
                    NotifyPropertyChanged("CurrentPage");
                }
            }
        }

        pdftron.PDF.PDFViewCtrl _myPDFViewCtrl = new PDFViewCtrl();
        pdftron.PDF.Tools.ToolManager _myToolManager;
        public pdftron.PDF.PDFViewCtrl MyPDFViewCtrl
        {
            get { return _myPDFViewCtrl; }
            set
            {
                if (_myPDFViewCtrl != value)
                {
                    _myPDFViewCtrl = value;
                    NotifyPropertyChanged("MyPDFViewCtrl");
                }
            }
        }

        public bool IsDocOpen
        {
            get { return MyPDFViewCtrl.HasDocument; }
        }

        private async void OpenFile()
        {
            FileOpenPicker fileOpenPicker = new FileOpenPicker();
            fileOpenPicker.ViewMode = PickerViewMode.List;
            fileOpenPicker.FileTypeFilter.Add(".pdf");
            StorageFile file = await fileOpenPicker.PickSingleFileAsync();
            if (file != null)
            {
                Windows.Storage.Streams.IRandomAccessStream stream = await
                    file.OpenAsync(FileAccessMode.ReadWrite);
                pdftron.PDF.PDFDoc doc = new pdftron.PDF.PDFDoc(stream);
                MyPDFViewCtrl.SetDoc(doc);            
                FileName = file.Name;
                PageCount = MyPDFViewCtrl.GetPageCount();
                ZoomAmount = 1;
                MyPDFViewCtrl.SetZoom(ZoomAmount);
            }
        }

        private async Task DragOpenFile(IStorageFile file)
        {
            if (file == null)
                return;
            Windows.Storage.Streams.IRandomAccessStream stream;
            try
            {
                stream = await file.OpenAsync(FileAccessMode.ReadWrite);
            }
            catch (Exception e)
            {
                // NOTE: If file already opened it will cause an exception
                var msg = new MessageDialog(e.Message);
                _ = msg.ShowAsync();
                return;
            }

            PDFDoc doc = new PDFDoc(stream);
            doc.InitSecurityHandler();

            // Set loaded doc to PDFView Controler 
            MyPDFViewCtrl.SetDoc(doc);
            FileName = file.Name;
            PageCount = MyPDFViewCtrl.GetPageCount();
            
            ZoomAmount = 1;
            MyPDFViewCtrl.SetZoom(ZoomAmount);
        }

        async public void ConvertFile()
        {
            FileOpenPicker fileOpenPicker = new FileOpenPicker();
            fileOpenPicker.ViewMode = PickerViewMode.List;
            fileOpenPicker.FileTypeFilter.Add(".docx");
            fileOpenPicker.FileTypeFilter.Add(".doc");
            fileOpenPicker.FileTypeFilter.Add(".pptx");
            StorageFile file = await fileOpenPicker.PickSingleFileAsync();
            if(file != null)
            {
                Windows.Storage.Streams.IRandomAccessStream stream;
                try
                {
                    stream = await file.OpenAsync(FileAccessMode.ReadWrite);
                }
                catch (Exception e)
                {
                    // NOTE: If file already opened it will cause an exception
                    var msg = new MessageDialog(e.Message);
                    _ = msg.ShowAsync();
                    return;
                }
                IFilter filter = new RandomAccessStreamFilter(stream);
                WordToPDFOptions opts = new WordToPDFOptions();
                DocumentConversion conversion = pdftron.PDF.Convert.UniversalConversion(filter, opts);
                var convRslt = conversion.TryConvert();

                if (convRslt == DocumentConversionResult.e_document_conversion_success)
                {
                    PDFDoc doc = conversion.GetDoc();
                    doc.InitSecurityHandler();

                    MyPDFViewCtrl.SetDoc(doc);
                }
            }
        }

        public void GoNextPage()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            MyPDFViewCtrl.GotoNextPage();
        }

        public void GoPrevPage()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            MyPDFViewCtrl.GotoPreviousPage();       
        }

        public void ZoomIn()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            ZoomAmount += .25;
            MyPDFViewCtrl.SetZoom(ZoomAmount);
        }

        public void ZoomOut()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            ZoomAmount -= .25;
            MyPDFViewCtrl.SetZoom(ZoomAmount);            
        }

        public void ResetZoom()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            ZoomAmount = 1;
            MyPDFViewCtrl.SetZoom(ZoomAmount);
        }

        public void CloseFile()
        {
            if (!MyPDFViewCtrl.HasDocument)
                return;
            MyPDFViewCtrl.CloseDoc();
            NotifyPropertyChanged("IsDocOpen");
        }

        public ICommand CMDOpenFile { get; set; }
        public ICommand CMDConvertFile { get; set; }
        public ICommand CMDGoNextPage { get; set; }
        public ICommand CMDGoPrevPage { get; set; }
        public ICommand CMDZoomIn { get; set; }
        public ICommand CMDZoomOut { get; set; }
        public ICommand CMDResetZoom { get; set; }
        public ICommand CMDCloseFile { get; set; }
    }
}
