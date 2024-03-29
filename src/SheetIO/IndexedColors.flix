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



mod SheetIO {


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

mod SheetIO.IndexedColors {

    use SheetIO.IndexedColors
    

    def toNative(x: IndexedColors): ##org.apache.poi.ss.usermodel.IndexedColors = 
        import static get org.apache.poi.ss.usermodel.IndexedColors.AQUA: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getAQUA;
        import static get org.apache.poi.ss.usermodel.IndexedColors.AUTOMATIC: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getAUTOMATIC;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BLACK: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBLACK;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BLACK1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBLACK1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BLUE_GREY: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBLUE_GREY;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BLUE1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBLUE1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BRIGHT_GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBRIGHT_GREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BRIGHT_GREEN1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBRIGHT_GREEN1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.BROWN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getBROWN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.CORAL: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getCORAL;
        import static get org.apache.poi.ss.usermodel.IndexedColors.CORNFLOWER_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getCORNFLOWER_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.DARK_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getDARK_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.DARK_GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getDARK_GREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.DARK_RED: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getDARK_RED;
        import static get org.apache.poi.ss.usermodel.IndexedColors.DARK_TEAL: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getDARK_TEAL;
        import static get org.apache.poi.ss.usermodel.IndexedColors.DARK_YELLOW: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getDARK_YELLOW;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GOLD: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGOLD;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGREY_25_PERCENT;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GREY_40_PERCENT: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGREY_40_PERCENT;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GREY_50_PERCENT: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGREY_50_PERCENT;
        import static get org.apache.poi.ss.usermodel.IndexedColors.GREY_80_PERCENT: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getGREY_80_PERCENT;
        import static get org.apache.poi.ss.usermodel.IndexedColors.INDIGO: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getINDIGO;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LAVENDER: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLAVENDER;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LEMON_CHIFFON: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLEMON_CHIFFON;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_CORNFLOWER_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_CORNFLOWER_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_GREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_ORANGE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_ORANGE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_TURQUOISE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_TURQUOISE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_TURQUOISE1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_TURQUOISE1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIGHT_YELLOW: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIGHT_YELLOW;
        import static get org.apache.poi.ss.usermodel.IndexedColors.LIME: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getLIME;
        import static get org.apache.poi.ss.usermodel.IndexedColors.MAROON: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getMAROON;
        import static get org.apache.poi.ss.usermodel.IndexedColors.OLIVE_GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getOLIVE_GREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.ORANGE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getORANGE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.ORCHID: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getORCHID;
        import static get org.apache.poi.ss.usermodel.IndexedColors.PALE_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getPALE_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.PINK: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getPINK;
        import static get org.apache.poi.ss.usermodel.IndexedColors.PINK1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getPINK1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.PLUM: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getPLUM;
        import static get org.apache.poi.ss.usermodel.IndexedColors.RED: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getRED;
        import static get org.apache.poi.ss.usermodel.IndexedColors.RED1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getRED1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.ROSE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getROSE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.ROYAL_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getROYAL_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.SEA_GREEN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getSEA_GREEN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.SKY_BLUE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getSKY_BLUE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.TAN: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getTAN;
        import static get org.apache.poi.ss.usermodel.IndexedColors.TEAL: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getTEAL;
        import static get org.apache.poi.ss.usermodel.IndexedColors.TURQUOISE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getTURQUOISE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.TURQUOISE1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getTURQUOISE1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.VIOLET: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getVIOLET;
        import static get org.apache.poi.ss.usermodel.IndexedColors.WHITE: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getWHITE;
        import static get org.apache.poi.ss.usermodel.IndexedColors.WHITE1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getWHITE1;
        import static get org.apache.poi.ss.usermodel.IndexedColors.YELLOW: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getYELLOW;
        import static get org.apache.poi.ss.usermodel.IndexedColors.YELLOW1: ##org.apache.poi.ss.usermodel.IndexedColors \ {} as getYELLOW1;
        match x {
            case SheetIO.IndexedColors.Aqua                 => getAQUA()
            case SheetIO.IndexedColors.Automatic            => getAUTOMATIC()
            case SheetIO.IndexedColors.Black                => getBLACK()
            case SheetIO.IndexedColors.Black1               => getBLACK1()
            case SheetIO.IndexedColors.Blue                 => getBLUE()
            case SheetIO.IndexedColors.BlueGrey             => getBLUE_GREY()
            case SheetIO.IndexedColors.Blue1                => getBLUE1()
            case SheetIO.IndexedColors.BrightGreen          => getBRIGHT_GREEN()
            case SheetIO.IndexedColors.BrightGreen1         => getBRIGHT_GREEN1()
            case SheetIO.IndexedColors.Brown                => getBROWN()
            case SheetIO.IndexedColors.Coral                => getCORAL()
            case SheetIO.IndexedColors.CornflowerBlue       => getCORNFLOWER_BLUE()
            case SheetIO.IndexedColors.DarkBlue             => getDARK_BLUE()
            case SheetIO.IndexedColors.DarkGreen            => getDARK_GREEN()
            case SheetIO.IndexedColors.DarkRed              => getDARK_RED()
            case SheetIO.IndexedColors.DarkTeal             => getDARK_TEAL()
            case SheetIO.IndexedColors.DarkYellow           => getDARK_YELLOW()
            case SheetIO.IndexedColors.Gold                 => getGOLD()
            case SheetIO.IndexedColors.Green                => getGREEN()
            case SheetIO.IndexedColors.Grey25Percent        => getGREY_25_PERCENT()
            case SheetIO.IndexedColors.Grey40Percent        => getGREY_40_PERCENT()
            case SheetIO.IndexedColors.Grey50Percent        => getGREY_50_PERCENT()
            case SheetIO.IndexedColors.Grey80Percent        => getGREY_80_PERCENT()
            case SheetIO.IndexedColors.Indigo               => getINDIGO()
            case SheetIO.IndexedColors.Lavender             => getLAVENDER()
            case SheetIO.IndexedColors.LemonChiffon         => getLEMON_CHIFFON()
            case SheetIO.IndexedColors.LightBlue            => getLIGHT_BLUE()
            case SheetIO.IndexedColors.LightCornflowerBlue  => getLIGHT_CORNFLOWER_BLUE() 
            case SheetIO.IndexedColors.LightGreen           => getLIGHT_GREEN()
            case SheetIO.IndexedColors.LightOrange          => getLIGHT_ORANGE()
            case SheetIO.IndexedColors.LightTurquoise       => getLIGHT_TURQUOISE()
            case SheetIO.IndexedColors.LightTurquoise1      => getLIGHT_TURQUOISE1()
            case SheetIO.IndexedColors.LightYellow          => getLIGHT_YELLOW()
            case SheetIO.IndexedColors.Lime                 => getLIME()
            case SheetIO.IndexedColors.Maroon               => getMAROON()
            case SheetIO.IndexedColors.OliveGreen           => getOLIVE_GREEN()
            case SheetIO.IndexedColors.Orange               => getORANGE()
            case SheetIO.IndexedColors.Orchid               => getORCHID()
            case SheetIO.IndexedColors.PaleBlue             => getPALE_BLUE()
            case SheetIO.IndexedColors.Pink                 => getPINK()
            case SheetIO.IndexedColors.Pink1                => getPINK1()
            case SheetIO.IndexedColors.Plum                 => getPLUM()
            case SheetIO.IndexedColors.Red                  => getRED()
            case SheetIO.IndexedColors.Red1                 => getRED1() 
            case SheetIO.IndexedColors.Rose                 => getROSE()
            case SheetIO.IndexedColors.RoyalBlue            => getROYAL_BLUE()
            case SheetIO.IndexedColors.SeaGreen             => getSEA_GREEN()
            case SheetIO.IndexedColors.SkyBlue              => getSKY_BLUE()
            case SheetIO.IndexedColors.Tan                  => getTAN()
            case SheetIO.IndexedColors.Teal                 => getTEAL()
            case SheetIO.IndexedColors.Turquoise            => getTURQUOISE()
            case SheetIO.IndexedColors.Turquoise1           => getTURQUOISE1()
            case SheetIO.IndexedColors.Violet               => getVIOLET()
            case SheetIO.IndexedColors.White                => getWHITE()
            case SheetIO.IndexedColors.White1               => getWHITE1()
            case SheetIO.IndexedColors.Yellow               => getYELLOW()
            case SheetIO.IndexedColors.Yellow1              => getYELLOW1()
        }

    pub def getIndex(x: IndexedColors): Int16 = 
        import org.apache.poi.ss.usermodel.IndexedColors.getIndex(): Int16 \ {};
        let x1 = toNative(x);
        getIndex(x1)

}
