using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Apuracao
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("ATENÇÃO!!");
            Console.WriteLine("Este programa gera apurações em excel referentes a ICMS, PIS e COFINS");
            Console.WriteLine("A partir do C100 e seus e registros referentes\n");
            Console.WriteLine("Insira o caminho para a pasta onde estão os arquivos SPED Fiscal");
            var root = Console.ReadLine();

            
            while (!Directory.Exists(root))
            {
                Console.WriteLine("Este diretório não é válido");
                root = Console.ReadLine();
            }

            Console.WriteLine("Aguarde...");

            var arquivos = Directory.GetFiles(root, "*.txt").ToList();

            arquivos.ForEach(arquivo =>
            {
                var arquivoCompleto = File.ReadAllLines(arquivo, Encoding.Default).ToList();

                var registros = arquivoCompleto.LeRegistros();

                var novoArquivo = Path.Combine(Path.GetDirectoryName(arquivo), Path.GetFileNameWithoutExtension(arquivo));


                var consolidacaoIcmsPorTipoRegistro = ConsolidacaoIcmsPorTipoRegistro.Gerar(registros);

                var consolidacaoIcmsPorCfop = ConsolidacaoIcmsPorCfop.Gerar(registros);

                var consolidacaoPisCofins = ConsolidacaoPisCofins.Gerar(registros);


                using (var excel = new ExcelPackage())
                {
                    var icmsRegistroCfop = excel.Workbook.Worksheets.Add("CONSOLIDADO_ICMS_REGISTRO_CFOP");

                    var linha = 1;
                    icmsRegistroCfop.Cells[linha, 2].Value = "Código";
                    icmsRegistroCfop.Cells[linha, 3].Value = "CFOP";
                    icmsRegistroCfop.Cells[linha, 4].Value = "Valor Contábil";
                    icmsRegistroCfop.Cells[linha, 5].Value = "Base de Cálculo";
                    icmsRegistroCfop.Cells[linha, 6].Value = "Valor de Icms";
                    consolidacaoIcmsPorTipoRegistro.ForEach(item =>
                    {
                        linha++;
                        icmsRegistroCfop.Cells[linha, 2].Value = item.Codigo;
                        icmsRegistroCfop.Cells[linha, 3].Value = item.Cfop;
                        icmsRegistroCfop.Cells[linha, 4].Value = item.ValorContabil;
                        icmsRegistroCfop.Cells[linha, 5].Value = item.BaseCalculoIcms;
                        icmsRegistroCfop.Cells[linha, 6].Value = item.Icms;
                    });

                    icmsRegistroCfop.Column(4).Style.Numberformat.Format = "#,##0.00";
                    icmsRegistroCfop.Column(5).Style.Numberformat.Format = "#,##0.00";
                    icmsRegistroCfop.Column(6).Style.Numberformat.Format = "#,##0.00";

                    icmsRegistroCfop.Row(1).Style.Font.Bold = true;
                    icmsRegistroCfop.Cells.AutoFitColumns(0);

                    var icmsCfop = excel.Workbook.Worksheets.Add("CONSOLIDADO_ICMS_CFOP");

                    linha = 1;
                    icmsCfop.Cells[linha, 2].Value = "CFOP";
                    icmsCfop.Cells[linha, 3].Value = "Valor Contábil";
                    icmsCfop.Cells[linha, 4].Value = "Base de Cálculo";
                    icmsCfop.Cells[linha, 5].Value = "Valor de Icms";

                    consolidacaoIcmsPorCfop.ForEach(item =>
                    {
                        linha++;
                        icmsCfop.Cells[linha, 2].Value = item.Cfop;
                        icmsCfop.Cells[linha, 3].Value = item.ValorContabil;
                        icmsCfop.Cells[linha, 4].Value = item.BaseCalculoIcms;
                        icmsCfop.Cells[linha, 5].Value = item.Icms;
                    });

                    icmsCfop.Column(3).Style.Numberformat.Format = "#,##0.00";
                    icmsCfop.Column(4).Style.Numberformat.Format = "#,##0.00";
                    icmsCfop.Column(5).Style.Numberformat.Format = "#,##0.00";

                    icmsCfop.Row(1).Style.Font.Bold = true;
                    icmsCfop.Cells.AutoFitColumns(0);


                    var pisCofinsCfop = excel.Workbook.Worksheets.Add("CONSOLIDADO_PIS_COFINS_CFOP");

                    linha = 1;
                    pisCofinsCfop.Cells[linha, 2].Value = "CFOP";
                    pisCofinsCfop.Cells[linha, 3].Value = "Base de Cálculo";
                    pisCofinsCfop.Cells[linha, 4].Value = "PIS";
                    pisCofinsCfop.Cells[linha, 5].Value = "COFINS";

                    consolidacaoPisCofins.ForEach(item =>
                    {
                        linha++;
                        pisCofinsCfop.Cells[linha, 2].Value = item.Cfop;
                        pisCofinsCfop.Cells[linha, 3].Value = item.BaseCalculoPisCofins;
                        pisCofinsCfop.Cells[linha, 4].Value = item.Pis;
                        pisCofinsCfop.Cells[linha, 5].Value = item.Cofins;

                    });

                    pisCofinsCfop.Column(3).Style.Numberformat.Format = "#,##0.00";
                    pisCofinsCfop.Column(4).Style.Numberformat.Format = "#,##0.00";
                    pisCofinsCfop.Column(5).Style.Numberformat.Format = "#,##0.00";

                    pisCofinsCfop.Row(1).Style.Font.Bold = true;
                    pisCofinsCfop.Cells.AutoFitColumns(0);

                    excel.SaveAs(new FileInfo(novoArquivo + "_apuracao.xlsx"));
                }

            });

            Console.WriteLine("Finalizado!");
            Console.WriteLine("Aperte em qualquer tecla para fechar esta janela");
            Console.ReadKey();
            Process.Start(root);
        }

        private static List<Registro> LeRegistros(this List<string> linhas)
        {
            var pula = true;

            var registros = new List<Registro>();

            linhas.ForEach(linha =>
            { 
                var data = linha.Split('|');

                if (pula && data[1].Equals("C100") && !data[6].Equals("90"))
                    pula = false;
                else
                {
                    var registro = Registro.Gerar(linha);
                    if (registro != null)
                        registros.Add(registro);

                    pula = true;
                }

            });

            return registros;

        }

    }

    public class ConsolidacaoIcmsPorCfop : Registro
    {
        public static List<ConsolidacaoIcmsPorCfop> Gerar(List<Registro> registros)
        {
            var listaConsolidacaoPorCfop = new List<ConsolidacaoIcmsPorCfop>();
            var cfops = registros.Select(registro => registro.Cfop).ToHashSet().ToList();

            cfops.ForEach(cfop =>
            {
                listaConsolidacaoPorCfop.Add(new ConsolidacaoIcmsPorCfop
                {
                    Cfop = cfop,
                    BaseCalculoIcms = registros
                                         .Where(registro => registro.Cfop.Equals(cfop))
                                         .Sum(registro => registro.BaseCalculoIcms),
                    ValorContabil = registros
                                         .Where(registro => registro.Cfop.Equals(cfop))
                                         .Sum(registro => registro.ValorContabil),
                    Icms = registros
                                    .Where(registro => registro.Cfop.Equals(cfop))
                                    .Sum(registro => registro.Icms),
                });
            });

            return listaConsolidacaoPorCfop;
        }

    }

    public class ConsolidacaoIcmsPorTipoRegistro : Registro
    {
        public static List<ConsolidacaoIcmsPorTipoRegistro> Gerar(List<Registro> registros)
        {
            var listaConsolidacaoPorRegistro = new List<ConsolidacaoIcmsPorTipoRegistro>();
            var cfops = registros.Select(registro => registro.Cfop).ToHashSet().ToList();
            var codigosRegistros = registros.Select(registro => registro.Codigo).ToHashSet().ToList();

            codigosRegistros.ForEach(codigo =>
            {
                cfops.ForEach(cfop =>
                {
                    listaConsolidacaoPorRegistro.Add(new ConsolidacaoIcmsPorTipoRegistro
                    {
                        Codigo = codigo,
                        Cfop = cfop,
                        BaseCalculoIcms = registros
                                             .Where(registro => registro.Cfop.Equals(cfop) && registro.Codigo.Equals(codigo))
                                             .Sum(registro => registro.BaseCalculoIcms),
                        ValorContabil = registros
                                             .Where(registro => registro.Cfop.Equals(cfop) && registro.Codigo.Equals(codigo))
                                             .Sum(registro => registro.ValorContabil),
                        Icms = registros
                                        .Where(registro => registro.Cfop.Equals(cfop) && registro.Codigo.Equals(codigo))
                                        .Sum(registro => registro.Icms),
                    });
                });

            });



            return listaConsolidacaoPorRegistro;
        }

        public override string ToString()
            => $";{Codigo};{Cfop};{ValorContabil.ToString("F2")};{BaseCalculoIcms.ToString("F2")};{Icms.ToString("F2")};";
    }

    public class ConsolidacaoPisCofins : Registro
    {
        public static List<ConsolidacaoPisCofins> Gerar(List<Registro> registros)
        {
            var listaConsolidacaoPorCfop = new List<ConsolidacaoPisCofins>();
            var cfops = registros.Select(registro => registro.Cfop).ToHashSet().ToList();

            cfops.ForEach(cfop =>
            {
                listaConsolidacaoPorCfop.Add(new ConsolidacaoPisCofins
                {
                    Cfop = cfop,


                    BaseCalculoPisCofins = registros
                                         .Where(registro => registro.Cfop.Equals(cfop))
                                         .Sum(registro => registro.BaseCalculoPisCofins),

                    Pis = registros
                                         .Where(registro => registro.Cfop.Equals(cfop))
                                         .Sum(registro => registro.Pis),

                    Cofins = registros
                                         .Where(registro => registro.Cfop.Equals(cfop))
                                         .Sum(registro => registro.Cofins)
                });
            });

            return listaConsolidacaoPorCfop;
        }
       
    }

    public class Registro
    {
        public static Registro Gerar(string linha)
        {
            try
            {
                var data = linha.Split('|');
                switch (data[1])
                {
                    case "C170":
                        return C170(data);
                    case "C190":
                        return C190(data);

                    default: return null;
                }
            }
            catch
            {

                Console.WriteLine("Corrija o problema na linha:\n\n" + linha);
                Console.WriteLine("\n\nEm seguida execute o porgraa novamente");
                Console.WriteLine("Aperte em qualquer tecla para fechar esta janela");
                Console.ReadKey();
                Environment.Exit(-1);
                throw;
            }
        }

        private static Registro C170(string[] data)
            => new Registro
            {
                Codigo = data[1],
                Item = data[3],
                ValorContabil = decimal.Parse(data[7]),
                Cfop = data[11],
                BaseCalculoIcms = decimal.Parse(data[13]),
                Aliquota = decimal.Parse(data[14]),
                Icms = decimal.Parse(data[15]),
                BaseCalculoPisCofins = decimal.Parse(data[26]),
                Pis = decimal.Parse(data[30]),
                Cofins = decimal.Parse(data[36])
            };

        private static Registro C190(string[] data)
            => new Registro
            {
                Codigo = data[1],
                Item = data[2],
                Cfop = data[3],
                Aliquota = decimal.Parse(data[4]),
                ValorContabil = decimal.Parse(data[5]),
                BaseCalculoIcms = decimal.Parse(data[6]),
                Icms = decimal.Parse(data[7])
            };



        public string Codigo { get; set; }
        public string Item { get; set; }
        public string Cfop { get; set; }
        public decimal ValorContabil { get; set; }
        public decimal BaseCalculoIcms { get; set; }
        public decimal Aliquota { get; set; }
        public decimal Icms { get; set; }
        public decimal BaseCalculoPisCofins { get; set; }
        public decimal Pis { get; set; }
        public decimal Cofins { get; set; }

    }
}
