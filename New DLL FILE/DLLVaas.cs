using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;

namespace New_DLL_FILE
{
    public class DLLVaas
    {
        private string alterPisorValor = "", lastPiso = "", dateHistorica, dateVigencia, newDir;
        private String[,] function;
        XLWorkbook wb, pisoRata;
        public string setNewTemplate(string dirNew, string bacaDir)
        {
            pisoRata = new XLWorkbook(dirNew);
            wb =  new XLWorkbook(bacaDir);
            newDir = @"S:\Gestao_Gente\Area_de_Servicos\Remuneração\1.Acordos Coletivos\Acordos Coletivos_BKP\2.Aplicações Acordo_2022\Robo Piso Salarial e Pro Rata\02. Template - Saida_" + DateTime.Today.ToString().Replace("/", "_").Replace("00:00:00", "") + ".xlsx";
            pisoRata.SaveAs(newDir);
            return newDir;
        }
        public string Init(List<String> listaObj, string igualOuMaior, string ateOuIgual, string piso, string valorPiso, string fixo, string valorFixo, int sindicato, int categoria, string municipio, string dataBase, string percentual, string proRata, string pisoIngresso, string minimoGarantido)
        {
            var pisoR = pisoRata.Worksheet(1);
            var aprendiz = pisoRata.Worksheet(2);
            var linhaPiso = pisoRata.Worksheet(1).Rows().Count();
            var linhaAprendiz = pisoRata.Worksheet(2).Rows().Count();
            var linhaExcluidos = pisoRata.Worksheet(3).Rows().Count();

            foreach (var item in listaObj)
            {

                var sindicatoBaca = wb.Worksheet(1).Cell("R" + item.Replace("R", "").ToString()).Value.ToString();
                var categoriaBaca = wb.Worksheet(1).Cell("T" + item.Replace("R", "").ToString()).Value.ToString();
                var cidade = wb.Worksheet(1).Cell("P" + item.Replace("R", "").ToString()).Value.ToString();

                if (sindicato == Int32.Parse(sindicatoBaca) && categoria == Int32.Parse(categoriaBaca) && municipio.Contains(cidade))
                {
                    var matricula = wb.Worksheet(1).Cell("C" + item.Replace("R", "").ToString()).Value.ToString();
                    var salario = wb.Worksheet(1).Cell("W" + item.Replace("R", "").ToString()).Value.ToString();
                    var jornarda = wb.Worksheet(1).Cell("V" + item.Replace("R", "").ToString()).Value.ToString();
                    var nome = wb.Worksheet(1).Cell("F" + item.Replace("R", "").ToString()).Value.ToString();
                    var empresa = wb.Worksheet(1).Cell("B" + item.Replace("R", "").ToString()).Value.ToString();
                    var uf = wb.Worksheet(1).Cell("O" + item.Replace("R", "").ToString()).Value.ToString();
                    var crm = wb.Worksheet(1).Cell("K" + item.Replace("R", "").ToString()).Value.ToString();
                    var estabelecimento = wb.Worksheet(1).Cell("M" + item.Replace("R", "").ToString()).Value.ToString();
                    var admisao = wb.Worksheet(1).Cell("D" + item.Replace("R", "").ToString()).Value.ToString();
                    var descricao = wb.Worksheet(1).Cell("S" + item.Replace("R", "").ToString()).Value.ToString();
                    var acordo = wb.Worksheet(1).Cell("Q" + item.Replace("R", "").ToString()).Value.ToString();
                    var codigoFuncao = wb.Worksheet(1).Cell("I" + item.Replace("R", "").ToString()).Value.ToString();
                    var funcao = wb.Worksheet(1).Cell("J" + item.Replace("R", "").ToString()).Value.ToString();
                    var funcaoCompa = wb.Worksheet(1).Cell("AA" + item.Replace("R", "").ToString()).Value.ToString().ToLower();

                    getPisoFunction(funcao, sindicato.ToString(), categoria.ToString(), salario.ToString());

                    if (funcao.Contains("Aprendiz".ToUpper()))
                    {
                        setEmployee(linhaAprendiz.ToString(), "Aprendiz - Excluido".ToString(),
                            matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());

                        aprendiz.Cell("AH" + linhaAprendiz).Value = "G_" + sindicato + "_" + cidade + "_APRENDIZ";
                        linhaAprendiz++;
                    }

                    else if (alterPisorValor != "")
                    {

                        if (Double.Parse(salario.ToString()) == Double.Parse(lastPiso.ToString()))
                        {
                            linhaPiso++;
                            setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                            pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL";
                            pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_PISO_POR_FUNCAO_" + alterPisorValor;
                            pisoR.Cell("AC" + linhaPiso).Value = Double.Parse(alterPisorValor.ToString()).ToString("F");
                        }
                    }
                    else if (piso.Contains("SIM"))
                    {
                        if ((Decimal.Parse(salario.ToString()) >= Decimal.Parse(igualOuMaior.ToString())) && (Decimal.Parse(salario.ToString()) <= Decimal.Parse(ateOuIgual.ToString())))
                        {
                            if (alterPisorValor != "")
                            {
                                if (Double.Parse(salario.ToString()) == Double.Parse(lastPiso.ToString()))
                                {
                                    linhaPiso++;
                                    setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                        matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                                    pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL";
                                    pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_PISO_POR_FUNCAO_" + alterPisorValor;
                                    pisoR.Cell("AC" + linhaPiso).Value = Double.Parse(alterPisorValor.ToString()).ToString("F");
                                    //return lastPiso ;
                                }
                            }
                            else
                            {
                                linhaPiso++;
                                setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                    matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                                pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL";
                                pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_ACIMA_" + igualOuMaior + "_MENOR_QUE_" + ateOuIgual;
                                pisoR.Cell("AC" + linhaPiso).Value = Double.Parse(valorPiso.ToString()).ToString("F");
                            }

                        }
                    }
                    else if (fixo.Contains("SIM") && (Decimal.Parse(salario.ToString()) >= Decimal.Parse(igualOuMaior.ToString())) && (Decimal.Parse(salario.ToString()) <= Decimal.Parse(ateOuIgual.ToString())))
                    {
                        linhaPiso++;
                        setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                        pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade;
                        pisoR.Cell("AC" + linhaPiso).Value = Double.Parse(valorFixo).ToString("F");
                        pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_FIXO_" + valorFixo;
                        pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL ";
                        pisoRata.Save();

                    }
                    else if (pisoIngresso == "SIM" && (Decimal.Parse(salario.ToString()) >= Decimal.Parse(igualOuMaior.ToString())) && (Decimal.Parse(salario.ToString()) <= Decimal.Parse(ateOuIgual.ToString())))
                    {
                        if (Decimal.Parse(salario.ToString()) == Decimal.Parse(valorPiso.ToString()))
                        {
                            linhaExcluidos++;
                            setEmployee(linhaExcluidos.ToString(), "Excluidos".ToString(),
                                matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                            pisoRata.Worksheet("Excluidos").Cell("AH" + linhaExcluidos).Value = "G_" + sindicato + "_" + cidade + "EXCLUIDO_SALARIO_IGUAL_PISO_INGRESO";
                        }
                        else
                        {

                                linhaPiso++;
                                setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                    matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                                pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL";
                                pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_ACIMA_" + valorPiso;

                        }

                    }
                    else if ((Decimal.Parse(salario.ToString()) >= Decimal.Parse(igualOuMaior.ToString())) && (Decimal.Parse(salario.ToString()) <= Decimal.Parse(ateOuIgual.ToString())) && pisoIngresso == "")
                    {
                        int result = DateTime.Compare(DateTime.Parse(admisao), DateTime.Parse(dataBase));
                        if (proRata == "SIM")
                        {
                            if (result >= 0)
                            {
                                linhaExcluidos++;
                                setEmployee(linhaExcluidos.ToString(), "Excluidos".ToString(),
                                    matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario.ToString());
                                pisoRata.Worksheet("Excluidos").Cell("AH" + linhaExcluidos).Value = "G_" + sindicato + "_" + cidade + "_EXCLUIDO_ADMISAO_APÓS";
                            }
                            else
                            {
                                linhaPiso++;
                                setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                    matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario);

                                pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_ACIMA_" + igualOuMaior;

                                admisao = admisao.Replace(" 00:00:00", "").Replace(" ", "");
                                dataBase = dataBase.Replace(" ", "");
                                string MonthCount = getPerc(admisao, dataBase, percentual);
                                pisoR.Cell("AE" + linhaPiso).Value = MonthCount + "%";
                                pisoR.Cell("U" + linhaPiso).Value = "PRO RATA";

                            }
                        }
                        else
                        {
                            linhaPiso++;
                            setEmployee(linhaPiso.ToString(), "Piso".ToString(),
                                matricula, nome, empresa, cidade, uf, crm, estabelecimento, admisao, descricao, acordo, categoria, codigoFuncao, funcao, sindicato, dataBase, jornarda, salario);

                            pisoR.Cell("AH" + linhaPiso).Value = "G_" + sindicato + "_" + cidade + "_ACIMA_" + igualOuMaior;

                            admisao = admisao.Replace(" 00:00:00", "").Replace(" ", "");
                            dataBase = dataBase.Replace(" ", "");
                            string MonthCount = getPerc(admisao, dataBase, percentual);
                            pisoR.Cell("AE" + linhaPiso).Value = percentual;
                            pisoR.Cell("U" + linhaPiso).Value = "INTEGRAL";
                           
                        }
                    }
                }
            }
            pisoRata.Save();
            return "Success";
        }
        public void setPisoFunction(String[,] functionExcel)
        {
            function = functionExcel;
        }
        public void setDate(string historica, string vigencia)
        {
            dateHistorica = historica;
            dateVigencia = vigencia;
        }
        private void setEmployee(string line, string sheet,
             string matricula, string nome, string empresa, string cidade, string uf, string crm, string estabelecimento, string admisao, string descricao,
             string acordo, int categoria, string codigoFuncao, string funcao, int sindicato, string dataBase, string jornarda, string salario)
        {
            pisoRata.Worksheet(sheet).Cell("A" + line).Value = matricula;
            pisoRata.Worksheet(sheet).Cell("O" + line).Value = matricula;
            pisoRata.Worksheet(sheet).Cell("C" + line).Value = nome;
            pisoRata.Worksheet(sheet).Cell("E" + line).Value = empresa;
            pisoRata.Worksheet(sheet).Cell("F" + line).Value = cidade;
            pisoRata.Worksheet(sheet).Cell("G" + line).Value = uf;
            pisoRata.Worksheet(sheet).Cell("H" + line).Value = crm;
            pisoRata.Worksheet(sheet).Cell("I" + line).Value = estabelecimento;
            pisoRata.Worksheet(sheet).Cell("N" + line).Value = admisao;
            pisoRata.Worksheet(sheet).Cell("P" + line).Value = descricao;
            pisoRata.Worksheet(sheet).Cell("Q" + line).Value = acordo;
            pisoRata.Worksheet(sheet).Cell("R" + line).Value = categoria;
            pisoRata.Worksheet(sheet).Cell("S" + line).Value = codigoFuncao;
            pisoRata.Worksheet(sheet).Cell("T" + line).Value = funcao;
            pisoRata.Worksheet(sheet).Cell("W" + line).Value = admisao;
            pisoRata.Worksheet(sheet).Cell("V" + line).Value = sindicato;
            pisoRata.Worksheet(sheet).Cell("X" + line).Value = dataBase;
            pisoRata.Worksheet(sheet).Cell("Y" + line).Value = dateHistorica;
            pisoRata.Worksheet(sheet).Cell("Z" + line).Value = dateVigencia;
            pisoRata.Worksheet(sheet).Cell("AA" + line).Value = jornarda; 
            pisoRata.Worksheet(sheet).Cell("AB" + line).SetValue(salario);
        }
        private string getPerc(string dateFuncionario, string datebase, string percentual)
        {
            string[] admisaoSplit = dateFuncionario.Split('/');
            string[] dataBaseSplit = datebase.Split('/');
            int MonthCount = ((12 * Int32.Parse(dataBaseSplit[2].ToString()) + Int32.Parse(dataBaseSplit[1])) - (12 * Int32.Parse(admisaoSplit[2].ToString()) + Int32.Parse(admisaoSplit[1])));
            String total = MonthCount.ToString();
            if (Int32.Parse(total) > 11) total = "12";
            else if (Int32.Parse(total) <= 0) total = "0";
            else if (Int32.Parse(admisaoSplit[0]) > 15)
            {
                total = (Int32.Parse(total) - 1).ToString();
            }

            Double perc = Double.Parse(percentual.Replace("%", "")) / 12 * Int32.Parse(total);
            return perc.ToString("N2");
        }
        private void getPisoFunction(string cargoFunc, string sindicato, string categoria, string salario)
        {
            string cargo = "";
            for (int i = 0; i < function.GetLength(0); i++)
            {
                cargo = function[i, 3].ToString();
                if (cargo.ToUpper() == cargoFunc.ToUpper() && function[i, 0] == sindicato && function[i, 1] == categoria && salario == function[i, 4])
                {
                    alterPisorValor = function[i, 5];
                    lastPiso = function[i, 4];
                    break;
                }
                else
                {
                    alterPisorValor = "";
                }
            }
        }
    }
}
