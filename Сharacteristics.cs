namespace PTS
{
    public struct Сharacteristics
    {
        public int number;
        public string name; //Наименование
        public double weight; //Масса 
        public double length; //Длина 
        public double width; //Ширина
        public double height; //Высота
        public double centerOfGravity; //Высота ЦТ
        public double Lpr; //Lпр
        public double Bp; //Bп
        public double L_CT; //Расстояние от поперечного борта вагона до Цт
        public double B_CT; //Расстояние от продольного борта вагона до Цт
        public double coefficientOfFriction; //коэф трения
        public double heightOfLongitudinal; //Высота продольного  упора от основания груза
        public double heightOfTransverse; //Высота поперчного упора от основания груза
        public double windwardSurfaceArea; //Площадь наветренной поверхности
        public double heightAboveFloor; //Высота груза над полом вагона
        public double HeightOfProtruding; //Высота выступающей части от бортов полувагона
        public double coefficientOfFrictionTransverse; //поперечный коэф трения
        public double additionalLongitudinalLoad; //Дополнительная продольная нагрузка от грузов №
        public double additionalLateralLoad; //Дополнительная поперечная нагрузка от грузов №
    }
}