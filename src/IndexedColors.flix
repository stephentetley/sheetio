/*
 * Copyright 2020 Stephen Tetley
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */



namespace SheetIO {


    pub enum IndexedColors { 
        case Aqua 
        case Automatic 
        case Black 
        case Black1 
        case Blue 
        case BlueGrey 
        case Blue1 
        case BrightGreen 
        case BrightGreen1 
        case Brown 
        case Coral 
        case CornflowerBlue 
        case DarkBlue 
        case DarkGreen 
        case DarkRed 
        case DarkTeal 
        case DarkYellow 
        case Gold 
        case Green 
        case Grey25Percent 
        case Grey40Percent 
        case Grey50Percent 
        case Grey80Percent 
        case Indigo 
        case Lavender 
        case LemonChiffon 
        case LightBlue 
        case LightCornflowerBlue 
        case LightGreen 
        case LightOrange 
        case LightTurquoise 
        case LightTurquoise1 
        case LightYellow 
        case Lime 
        case Maroon 
        case OliveGreen 
        case Orange 
        case Orchid 
        case PaleBlue 
        case Pink 
        case Pink1 
        case Plum 
        case Red 
        case Red1 
        case Rose 
        case RoyalBlue 
        case SeaGreen 
        case SkyBlue 
        case Tan 
        case Teal 
        case Turquoise 
        case Turquoise1 
        case Violet 
        case White 
        case White1 
        case Yellow 
        case Yellow1 
    }

}

namespace SheetIO/IndexedColors {

    use SheetIO.IndexedColors;
    

    def toNative(x: IndexedColors): ##org.apache.poi.ss.usermodel.IndexedColors = 
        import get org.apache.poi.ss.usermodel.IndexedColors:AQUA                   as getAQUA;
        import get org.apache.poi.ss.usermodel.IndexedColors:AUTOMATIC              as getAUTOMATIC;
        import get org.apache.poi.ss.usermodel.IndexedColors:BLACK                  as getBLACK;
        import get org.apache.poi.ss.usermodel.IndexedColors:BLACK1                 as getBLACK1;
        import get org.apache.poi.ss.usermodel.IndexedColors:BLUE                   as getBLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:BLUE_GREY              as getBLUE_GREY;
        import get org.apache.poi.ss.usermodel.IndexedColors:BLUE1                  as getBLUE1;
        import get org.apache.poi.ss.usermodel.IndexedColors:BRIGHT_GREEN           as getBRIGHT_GREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:BRIGHT_GREEN1          as getBRIGHT_GREEN1;
        import get org.apache.poi.ss.usermodel.IndexedColors:BROWN                  as getBROWN;
        import get org.apache.poi.ss.usermodel.IndexedColors:CORAL                  as getCORAL;
        import get org.apache.poi.ss.usermodel.IndexedColors:CORNFLOWER_BLUE        as getCORNFLOWER_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:DARK_BLUE              as getDARK_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:DARK_GREEN             as getDARK_GREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:DARK_RED               as getDARK_RED;
        import get org.apache.poi.ss.usermodel.IndexedColors:DARK_TEAL              as getDARK_TEAL;
        import get org.apache.poi.ss.usermodel.IndexedColors:DARK_YELLOW            as getDARK_YELLOW;
        import get org.apache.poi.ss.usermodel.IndexedColors:GOLD                   as getGOLD;
        import get org.apache.poi.ss.usermodel.IndexedColors:GREEN                  as getGREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:GREY_25_PERCENT        as getGREY_25_PERCENT;
        import get org.apache.poi.ss.usermodel.IndexedColors:GREY_40_PERCENT        as getGREY_40_PERCENT;
        import get org.apache.poi.ss.usermodel.IndexedColors:GREY_50_PERCENT        as getGREY_50_PERCENT;
        import get org.apache.poi.ss.usermodel.IndexedColors:GREY_80_PERCENT        as getGREY_80_PERCENT;
        import get org.apache.poi.ss.usermodel.IndexedColors:INDIGO                 as getINDIGO;
        import get org.apache.poi.ss.usermodel.IndexedColors:LAVENDER               as getLAVENDER;
        import get org.apache.poi.ss.usermodel.IndexedColors:LEMON_CHIFFON          as getLEMON_CHIFFON;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_BLUE             as getLIGHT_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_CORNFLOWER_BLUE  as getLIGHT_CORNFLOWER_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_GREEN            as getLIGHT_GREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_ORANGE           as getLIGHT_ORANGE;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_TURQUOISE        as getLIGHT_TURQUOISE;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_TURQUOISE1       as getLIGHT_TURQUOISE1;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIGHT_YELLOW           as getLIGHT_YELLOW;
        import get org.apache.poi.ss.usermodel.IndexedColors:LIME                   as getLIME;
        import get org.apache.poi.ss.usermodel.IndexedColors:MAROON                 as getMAROON;
        import get org.apache.poi.ss.usermodel.IndexedColors:OLIVE_GREEN            as getOLIVE_GREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:ORANGE                 as getORANGE;
        import get org.apache.poi.ss.usermodel.IndexedColors:ORCHID                 as getORCHID;
        import get org.apache.poi.ss.usermodel.IndexedColors:PALE_BLUE              as getPALE_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:PINK                   as getPINK;
        import get org.apache.poi.ss.usermodel.IndexedColors:PINK1                  as getPINK1;
        import get org.apache.poi.ss.usermodel.IndexedColors:PLUM                   as getPLUM;
        import get org.apache.poi.ss.usermodel.IndexedColors:RED                    as getRED;
        import get org.apache.poi.ss.usermodel.IndexedColors:RED1                   as getRED1;
        import get org.apache.poi.ss.usermodel.IndexedColors:ROSE                   as getROSE;
        import get org.apache.poi.ss.usermodel.IndexedColors:ROYAL_BLUE             as getROYAL_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:SEA_GREEN              as getSEA_GREEN;
        import get org.apache.poi.ss.usermodel.IndexedColors:SKY_BLUE               as getSKY_BLUE;
        import get org.apache.poi.ss.usermodel.IndexedColors:TAN                    as getTAN;
        import get org.apache.poi.ss.usermodel.IndexedColors:TEAL                   as getTEAL;
        import get org.apache.poi.ss.usermodel.IndexedColors:TURQUOISE              as getTURQUOISE;
        import get org.apache.poi.ss.usermodel.IndexedColors:TURQUOISE1             as getTURQUOISE1;
        import get org.apache.poi.ss.usermodel.IndexedColors:VIOLET                 as getVIOLET;
        import get org.apache.poi.ss.usermodel.IndexedColors:WHITE                  as getWHITE;
        import get org.apache.poi.ss.usermodel.IndexedColors:WHITE1                 as getWHITE1;
        import get org.apache.poi.ss.usermodel.IndexedColors:YELLOW                 as getYELLOW;
        import get org.apache.poi.ss.usermodel.IndexedColors:YELLOW1                as getYELLOW1;
        match x {
            case SheetIO/IndexedColors.Aqua                 => getAQUA()
            case SheetIO/IndexedColors.Automatic            => getAUTOMATIC()
            case SheetIO/IndexedColors.Black                => getBLACK()
            case SheetIO/IndexedColors.Black1               => getBLACK1()
            case SheetIO/IndexedColors.Blue                 => getBLUE()
            case SheetIO/IndexedColors.BlueGrey             => getBLUE_GREY()
            case SheetIO/IndexedColors.Blue1                => getBLUE1()
            case SheetIO/IndexedColors.BrightGreen          => getBRIGHT_GREEN()
            case SheetIO/IndexedColors.BrightGreen1         => getBRIGHT_GREEN1()
            case SheetIO/IndexedColors.Brown                => getBROWN()
            case SheetIO/IndexedColors.Coral                => getCORAL()
            case SheetIO/IndexedColors.CornflowerBlue       => getCORNFLOWER_BLUE()
            case SheetIO/IndexedColors.DarkBlue             => getDARK_BLUE()
            case SheetIO/IndexedColors.DarkGreen            => getDARK_GREEN()
            case SheetIO/IndexedColors.DarkRed              => getDARK_RED()
            case SheetIO/IndexedColors.DarkTeal             => getDARK_TEAL()
            case SheetIO/IndexedColors.DarkYellow           => getDARK_YELLOW()
            case SheetIO/IndexedColors.Gold                 => getGOLD()
            case SheetIO/IndexedColors.Green                => getGREEN()
            case SheetIO/IndexedColors.Grey25Percent        => getGREY_25_PERCENT()
            case SheetIO/IndexedColors.Grey40Percent        => getGREY_40_PERCENT()
            case SheetIO/IndexedColors.Grey50Percent        => getGREY_50_PERCENT()
            case SheetIO/IndexedColors.Grey80Percent        => getGREY_80_PERCENT()
            case SheetIO/IndexedColors.Indigo               => getINDIGO()
            case SheetIO/IndexedColors.Lavender             => getLAVENDER()
            case SheetIO/IndexedColors.LemonChiffon         => getLEMON_CHIFFON()
            case SheetIO/IndexedColors.LightBlue            => getLIGHT_BLUE()
            case SheetIO/IndexedColors.LightCornflowerBlue  => getLIGHT_CORNFLOWER_BLUE() 
            case SheetIO/IndexedColors.LightGreen           => getLIGHT_GREEN()
            case SheetIO/IndexedColors.LightOrange          => getLIGHT_ORANGE()
            case SheetIO/IndexedColors.LightTurquoise       => getLIGHT_TURQUOISE()
            case SheetIO/IndexedColors.LightTurquoise1      => getLIGHT_TURQUOISE1()
            case SheetIO/IndexedColors.LightYellow          => getLIGHT_YELLOW()
            case SheetIO/IndexedColors.Lime                 => getLIME()
            case SheetIO/IndexedColors.Maroon               => getMAROON()
            case SheetIO/IndexedColors.OliveGreen           => getOLIVE_GREEN()
            case SheetIO/IndexedColors.Orange               => getORANGE()
            case SheetIO/IndexedColors.Orchid               => getORCHID()
            case SheetIO/IndexedColors.PaleBlue             => getPALE_BLUE()
            case SheetIO/IndexedColors.Pink                 => getPINK()
            case SheetIO/IndexedColors.Pink1                => getPINK1()
            case SheetIO/IndexedColors.Plum                 => getPLUM()
            case SheetIO/IndexedColors.Red                  => getRED()
            case SheetIO/IndexedColors.Red1                 => getRED1() 
            case SheetIO/IndexedColors.Rose                 => getROSE()
            case SheetIO/IndexedColors.RoyalBlue            => getROYAL_BLUE()
            case SheetIO/IndexedColors.SeaGreen             => getSEA_GREEN()
            case SheetIO/IndexedColors.SkyBlue              => getSKY_BLUE()
            case SheetIO/IndexedColors.Tan                  => getTAN()
            case SheetIO/IndexedColors.Teal                 => getTEAL()
            case SheetIO/IndexedColors.Turquoise            => getTURQUOISE()
            case SheetIO/IndexedColors.Turquoise1           => getTURQUOISE1()
            case SheetIO/IndexedColors.Violet               => getVIOLET()
            case SheetIO/IndexedColors.White                => getWHITE()
            case SheetIO/IndexedColors.White1               => getWHITE1()
            case SheetIO/IndexedColors.Yellow               => getYELLOW()
            case SheetIO/IndexedColors.Yellow1              => getYELLOW1()
        } as & Pure

    pub def getIndex(x: IndexedColors): Int16 = 
        import org.apache.poi.ss.usermodel.IndexedColors.getIndex();
        let x1 = toNative(x);
        getIndex(x1) as & Pure


}