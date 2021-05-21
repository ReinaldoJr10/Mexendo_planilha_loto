using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Mexendo_planilha
{
    public class Sorteio
    {
        public int numeroSorteio { get; set; }
        public string data { get; set; }
        public int[] numeros { get; set; }
    }
    class Program
    {
        static int ContaPar(int[] numeros)
        {
            int acu = 0;
            var tam = numeros.Length;
            for(int i = 0; i < tam; i++)
            {
                if (numeros[i] % 2 == 0)
                    acu++;
            }
            return acu;
        }
        static int ContaImpar(int[] numeros)
        {
            int acu = 0;
            var tam = numeros.Length;
            for (int i = 0; i < tam; i++)
            {
                if (numeros[i] % 2 == 1)
                    acu++;
            }
            return acu;
        }
        static int ContaPrimo(int[] numeros)
        {
            int acu = 0,cont=0;
            for(int i = 0; i < numeros.Length; i++)
            {
                for(int j = 1; j <= numeros[i]; j++)
                {
                    if (numeros[i] % j == 0)
                        cont++;
                }
                if (cont == 2)
                    acu++;
                cont = 0;
            }
            return acu;            
        }
        static void Main(string[] args)
        {
            var dados = new List<Sorteio>();
            var xls = new XLWorkbook(@"C:\Users\Junior\Downloads\loto.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "lotofacil_www.asloterias.com.br");
            for (int l = planilha.Rows().Count() ; l !=2; l=l-1)
            {
                var aux = new Sorteio();
                int[] combinacao = new int[15];
                aux.numeroSorteio= int.Parse(planilha.Cell($"A{l}").Value.ToString());
                aux.data = planilha.Cell($"B{l}").Value.ToString();
                for(int i = 0; i < 15; i++)                
                    combinacao[i] = int.Parse(planilha.Cell($"{(char)(67 + i)}{l}").Value.ToString());                
                Array.Sort(combinacao);
                aux.numeros = combinacao;
                dados.Add(aux);
                
            }
            foreach(var temp in dados)            
                Console.WriteLine($"{temp.numeroSorteio} - {temp.data} \n"+ String.Join(",", temp.numeros)+ " Par: " + ContaPar(temp.numeros) + " Impar: " + ContaImpar(temp.numeros) +" Primos: "+ContaPrimo(temp.numeros)+"\n--------------------------------------------------------------------------\n");            
        }
    }
}
