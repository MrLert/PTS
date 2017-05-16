using System.Collections.Generic;

namespace PTS
{
    public class Calculations
    {
        public List<СarriageСharacteristic> CarriageList;
        public List<Сharacteristics> CharacteristicsList;
        public OfficeWord Word = new OfficeWord();

        public Calculations(List<СarriageСharacteristic> carriageСharacteristics, List<Сharacteristics> сharacteristics)
        {
            CarriageList = carriageСharacteristics;
            CharacteristicsList = сharacteristics;
        }

        public void LongitudinalHorizontalInertialForces()
        {
        }

        public void CrossHorizontalInertialForces()
        {
        }

        public void VerticalInertialForces()
        {
        }

        public void FrictionForces()
        {
        }
    }
}