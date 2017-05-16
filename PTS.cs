using System.Windows.Forms;

namespace PTS
{
    public partial class PTS : Form
    {
        public PTS()
        {
            InitializeComponent();
            var excel = new OfficeExcel();
            var infoTitle = excel.InpuTitle();
            var CarriageCharacteristic = excel.InputListCarriage();
            var ListCharacteristicses = excel.InputListСharacteristicses();
            var count = excel.count;
            excel.CloseExcel();
            var word = new WordMainReport(infoTitle, CarriageCharacteristic, ListCharacteristicses, count);
            word.createTitle();
        }
    }
}