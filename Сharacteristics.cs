namespace PTS
{
    struct Сharacteristics
    {
        string name;                                //Наименование
        double weight;                              //Масса 
        double length;                              //Длина 
        double width;                               //Ширина
        double height;                              //Высота
        double centerOfGravity;                     //Высота ЦТ
        double Lpr;                                 //Lпр
        double Bp;                                  //Bп
        double L_CT;                                //Расстояние от поперечного борта вагона до Цт
        double B_CT;                                //Расстояние от продольного борта вагона до Цт
        double coefficientOfFriction;               //коэф трения
        double heightOfLongitudinal;                //Высота продольного  упора от основания груза
        double heightOfTransverse;                  //Высота поперчного упора от основания груза
        double windwardSurfaceArea;                 //Площадь наветренной поверхности
        double heightAboveFloor;                    //Высота груза над полом вагона
        double HeightOfProtruding;                  //Высота выступающей части от бортов полувагона
        double coefficientOfFrictionTransverse;     //поперечный коэф трения
        double additionalLongitudinalLoad;          //Дополнительная продольная нагрузка от грузов №
        double additionalLateralLoad;               //Дополнительная поперечная нагрузка от грузов №
    }
}