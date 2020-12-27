using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace veremautomata
{
    class Program
    {
        static void Main(string[] args)
        {
            #region beolvasas
            //excel tábla megadása
            Application excelApp = new Application();
            Workbook excelBook = excelApp.Workbooks.Open(@"/Users/norbertsass-gyarmati/Desktop/verem_automata/verem_automata/verem_tabla.xlsx");//D:\verem_automata\verem_automata\verem_tabla.xlsx
            //"\Users\norbertsass-gyarmati\Desktop\verem_automata\verem_automata\verem_tabla.xlsx"
            _Worksheet excelSheet = (_Worksheet)excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            

            //változók megadása
            int rows = excelRange.Rows.Count; 
            int cols = excelRange.Columns.Count;
            string[,] tabla = new string[rows-1, cols-1]; //tényleges táblát létrehozom, majd ebbe fogom tudni az excel celláit tárolni
            List<string> rowSymbols = new List<string>(); // a sorok szimbólumait egy listában tárolom
            List<string> colSymbols = new List<string>(); //Ahogy az oszlopokat is

            //1. oszlop jeleinek meghatározása
            for (int i = 2; i < cols+1; i++)
            {
                colSymbols.Add(excelRange.Cells[1, i].ToString()); // feltöltöm az oszlopok szimbólumait
            }

            //1. sor jeleinek meghatározása
            for (int i = 2; i < rows+1; i++)
            {
                rowSymbols.Add(excelRange.Cells[i, 1].ToString()); //feltöltöm a sorok szimbólumait
            }
            
            //táblázat kiolvasása
            for (int i = 2; i <= rows; i++)
            {
                for (int j = 2; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null) //Ha tartalmaz valamit az adott cella, akkor betöltjük a táblába
                    {
                        tabla[i-2,j-2]=excelRange.Cells[i, j].ToString(); //Muszáj a tostring, másképp nem ismeri fel.
                    }
                    else
                    {
                        tabla[i - 2, j - 2] = "error"; // aműgy pedig egy error értéket töltünk be a tábla adott cellájába
                    }   
                }
            }

            //excel bezárása
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp); //Csökkenti a megadott excelApphoz társított futásidejű hívható csomagoló referenciaszámát.
            #endregion

            //kifejezés bekérése
            Console.WriteLine("Kérek egy kifejezést:");
            string kif = Console.ReadLine() + "#"; //az volt a feladat, hogy a végén legyen egy elfogadó jel #
            bool isTrue = true; //ha nem változik az értéke, akkor a lefutás után elfogadó állapotban lesz

            //stack előkészítése
            Stack<string> stack = new Stack<string>(); //stacket hozok létre
            stack.Push("#"); //első elem mindig a # jel
            stack.Push(rowSymbols[0]); //illetve belerakom még a rowSymbol legelső elemét

            //kiértékelés
            for (int i = 0; i < kif.Length;) 
            {
                string currentCellContent = tabla[rowSymbols.IndexOf(stack.Pop()), colSymbols.IndexOf(Convert.ToString(kif[i]))];
                //az aktuális cellának értékül adom a létrehozott tábla aktuális indexén lévő elemet
                if (currentCellContent == "pop") //ha ennek az értéke pop, akkor, akkor ugye el kell törölni azt az elemet, amit én úgy oldok meg, hogy egyszerűen továbbléptetem az input szalagot
                {
                    i++;
                }
                else if (currentCellContent == "e") { } //ilyen esetben nem történk semmi
                else if (currentCellContent=="accept") //ki kell ugorjunk a ciklusból, hogy as isTrue értéke ne változzon
                {
                    break;
                }
                else if (currentCellContent != "error") //ha nem kapunk errort, azaz minden cellának van értéke, akkor
                    //ismét elvágjuk az input szalagot és betöltjük az aktuális elfogadott elemet
                {
                    string[] currentCellSplit = currentCellContent.Split(' ');
                    for (int j = currentCellSplit.Length - 1; j >= 0; j--)
                    {
                        stack.Push(currentCellSplit[j]);
                    } //a kövi lefutásnál már nem lesz benne az előző eredmény, így a következőt tudja levizsgálni.
                }
                else //ha viszont més értéket kap, akkor az istrue hamis lesz így a kifejezés helytelen lesz
                {
                    isTrue = false;
                    break;
                }
            }

            //eredmény kiiratása
            if (isTrue)
            {
                Console.WriteLine("A kifejezés helyes!");
            }
            else
            {
                Console.WriteLine("A kifejezés helytelen!");
            }
            Console.ReadKey();
        }
    }
}


