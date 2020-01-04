using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace AlteradorDeRegistros
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Insira o caminho para a pasta onde estão os arquivos SPED Fiscal da FAL");
            var root = Console.ReadLine();

            while (!Directory.Exists(root))
            {
                Console.WriteLine("Este diretório não é válido");
                root = Console.ReadLine();
            }

            Console.WriteLine("Aguarde...");

            var arquivos = Directory.GetFiles(root,"*.txt").ToList();

            arquivos.ForEach(arquivo=> 
            {
                var arquivoCompleto = File.ReadAllLines(arquivo , Encoding.Default).ToList();

                var registros = arquivoCompleto.Select(linha => Registro.Gerar(linha)).ToList();

                registros.NomeFantasia();

                registros.DeletaObsoleto();

                registros.DeletaC170();

                registros.Delete0200();

                var novosRegistros = registros.Where(x => !x.Sai)
                                          .Select(x => x.Linha).ToList();


                var novoArquivo = Path.Combine(Path.GetDirectoryName(arquivo), Path.GetFileNameWithoutExtension(arquivo) + "_alterado.txt");

                File.WriteAllLines(novoArquivo, novosRegistros, Encoding.Default);
            });


            Console.WriteLine("Finalizado!");
            Console.WriteLine("Aperte em qualquer tecla para fechar esta janela");
            Console.ReadKey();
            Process.Start(root);

        }


        private static void NomeFantasia(this List<Registro> registros)
        {
            var registro0005 = registros.FirstOrDefault(x => x.Codigo.Equals("0005") && string.IsNullOrEmpty(x.Item));

            registro0005.Linha = registro0005.Linha.Replace("||", "|FAL - FÁBRICA DE ALIMENTOS|");
        }

        private static void DeletaObsoleto(this List<Registro> registros)
        {
            registros.RemoveAll(x => x.Codigo.Equals("H010")
                                 || x.Codigo.Equals("C177"));
        }



        private static void DeletaC170(this List<Registro> registros)
        {
            var delete = false;
            registros.ForEach(x =>
            {
                if (x.Codigo.Equals("C100"))
                {
                    if (x.Linha.Contains("|C100|0|0|") || x.Linha.Contains("|C100|1|0|"))
                    {
                        delete = true;
                    }
                    else
                    {
                        delete = false;
                    }
                }
                if (x.Codigo.Equals("C170") && delete)
                {
                    x.Sai = true;
                }
            });

        }

        private static void Delete0200(this List<Registro> registros)
        {
            var listaItensUtilizados = registros.Where(x => !x.Codigo.Equals("0200") && !x.Sai).Select(x => x.Item);

            registros.RemoveAll(x => x.Codigo.Equals("0200")
                                   && !listaItensUtilizados.Contains(x.Item));
        }

    }

    public class Registro
    {
        private Registro()
        {
        }

        public static Registro Gerar(string linha)
        {
            var data = linha.Split('|');

            return new Registro
            {
                Codigo = data[1],
                Item = GetItem(data),
                Linha = linha
            };

        }

        private static string GetItem(string[] data)
        {


            switch (data[1])
            {
                case "0200":
                case "H010":
                case "C590":
                case "D590":
                case "C321":
                case "C320":
                case "C425":
                    return data[2];

                case "C170":
                    return data[3];

                default: return "";
            }
        }

        public string Codigo { get; set; }
        public string Item { get; set; }
        public string Linha { get; set; }
        public bool Sai { get; set; }
    }
}
