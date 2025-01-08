using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WpfApp1.ViewModels
{
    public class MakeReservationViewModel : ViewModelBase
    {
        private string? _userName;
        public string UserName { get => _userName ?? "no username"; set { _userName = value; OnPropertyChanged(nameof(UserName)); } }

        private int _floorNumber;
        public int FloorNumber { get => _floorNumber; set { _floorNumber = value; OnPropertyChanged(nameof(FloorNumber)); } }

        private int _roomNumber;
        public int RoomNumber { get => _roomNumber; set { _roomNumber = value; OnPropertyChanged(nameof(RoomNumber)); } }

        private DateTime _startDate;
        public DateTime StartDate { get => _startDate; set { _startDate = value; OnPropertyChanged(nameof(StartDate)); } }

        private DateTime _endDate;
        public DateTime EndDate { get => _endDate; set { _endDate = value; OnPropertyChanged(nameof(EndDate)); } }

        public ICommand SubmitCommand { get; }
        public ICommand CancelCommand { get; }

        public MakeReservationViewModel()
        {
            
        }
    }
}
