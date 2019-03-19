using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Input;
using Microsoft.Win32;

namespace ExcelWriter
{
    class ExcelWriterViewController : INotifyPropertyChanged
    {
        private CancellationTokenSource cancellationTokenSource;

        public ExcelWriterData Data { get; set; }

        private string filename = string.Empty;
        public string Filename { get => filename; set { filename = value; NotifyPropertyChanged(); } }

        private bool isBusy = false;
        private bool isCanceling = false;

        public RelayCommand AppendCommand { get; set; }
        public RelayCommand CancelCommand { get; set; }
        public RelayCommand ClearCommand { get; set; }
        public RelayCommand SelectFileCommand { get; set; }

        public ExcelWriterViewController()
        {
            Data = new ExcelWriterData();
            AppendCommand = new RelayCommand(Append, IsIdle);
            CancelCommand = new RelayCommand(Cancel, CanCancel);
            ClearCommand = new RelayCommand(Clear, IsIdle);
            SelectFileCommand = new RelayCommand(SelectFile, IsIdle);

        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        bool IsIdle()
        {
            return !isBusy;
        }

        bool CanCancel()
        {
            return !isCanceling && isBusy;
        }

        void updateBusy(bool busyValue)
        {
            isBusy = busyValue;
            AppendCommand.OnCanExecuteChanged();
            CancelCommand.OnCanExecuteChanged();
            ClearCommand.OnCanExecuteChanged();
            SelectFileCommand.OnCanExecuteChanged();
        }

        public async void Append()
        {
            updateBusy(true);
            cancellationTokenSource = new CancellationTokenSource();
            await Data.AppendToFileAsync(filename, cancellationTokenSource.Token);
            updateBusy(false);
            isCanceling = false;
        }

        public void Cancel()
        {
            isCanceling = true;
            CancelCommand.OnCanExecuteChanged();
            Data.Result = "Appand operation is being canceled";
            if (cancellationTokenSource != null)
                cancellationTokenSource.Cancel();
        }

        public void Clear()
        {
            Data.ClearValues();
        }

        public void SelectFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dokumente (*.xlsx, *.xls, *.xlsm, *.xlst) | *.xlsx;*.xls;*.xlsm;*.xlst";
            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                Filename = openFileDialog.FileName;
            }
        }
    }
}
