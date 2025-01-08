using System.Collections.ObjectModel;
using System.Windows.Input;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{
    public class ReservationListingViewModel : ViewModelBase
    {
        private readonly ObservableCollection<ReservationViewModel> _reservations;
        public IEnumerable<ReservationViewModel> Reservations => _reservations;
        public ICommand MakeReservationCommand { get; }


        public ReservationListingViewModel()
        {
            _reservations = new ObservableCollection<ReservationViewModel>
            {
                new ReservationViewModel(new Reservation(new RoomID(0, 1), "Test1", DateTime.Now, DateTime.Now)),
                new ReservationViewModel(new Reservation(new RoomID(0, 2), "Test2", DateTime.Now, DateTime.Now)),
                new ReservationViewModel(new Reservation(new RoomID(1, 1), "Test3", DateTime.Now, DateTime.Now))
            };
        }
    }
}
