using System.Collections.Generic;

namespace PTS
{
    public class AddCalculations
    {
        public List<СarriageСharacteristic> CarriageList;
        public List<Сharacteristics> CharacteristicsList;
        public OfficeWord Word = new OfficeWord();

        public AddCalculations(List<СarriageСharacteristic> carriageСharacteristics,
            List<Сharacteristics> сharacteristics)
        {
            CarriageList = carriageСharacteristics;
            CharacteristicsList = сharacteristics;
        }

        public void NailFastenning()
        {
        }

        public void WireFastenning()
        {
        }
    }
}