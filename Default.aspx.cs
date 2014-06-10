using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

public partial class _Default : System.Web.UI.Page
{
    DataSet dsClassificados = new DataSet();
    DataSet dsOpcoes = new DataSet();
    DataSet dsVagas = new DataSet();
    DataSet dsEstudo = new DataSet();
    DataSet dsDesistencias = new DataSet();
    List<String> lsVagasAbertas = new List<string>();
    static int maxOpcoes = 241;
    static bool rodarDeNovo = true;
    static bool temDesistencias = true;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    private bool desistiu(int i, int j)
    {
        if (temDesistencias)
        {
            string classif = dsOpcoes.Tables[0].Rows[i][0].ToString();
            string codUA = dsOpcoes.Tables[0].Rows[i][j].ToString().Substring(0, 7).TrimStart('0');

            DataRow[] dr = dsDesistencias.Tables[0].Select("Class = '" + classif + "' AND CodUA = '" + codUA + "'");

            if (dr.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        else
        {
            return false;
        }
    }

    private void simular()
    {
        bool houveMudanca = true;
        while (houveMudanca)
        {
            houveMudanca = false;
            for (int i = 0; i < dsClassificados.Tables[0].Rows.Count; i++)
            {
                for (int j = 1; j <= maxOpcoes; j++)
                {
                    if (dsOpcoes.Tables[0].Rows[i][j] != DBNull.Value && (string)dsOpcoes.Tables[0].Rows[i][j] != string.Empty)
                    {

                        if (!desistiu(i,j))
                        {

                            DataRow[] dr = dsVagas.Tables[0].Select("Unidade = '" + dsOpcoes.Tables[0].Rows[i][j].ToString().Replace("'", "''") + "'");
                            if (dr.Count() > 0 && Convert.ToInt32(dr[0]["Vagas"]) > 0)
                            {
                                //marca que houve mudança
                                houveMudanca = true;

                                //coloca a lotação do candidato numa variável temporária
                                String tempLotacao = (string)dsClassificados.Tables[0].Rows[i]["Lotação Final"];

                                //muda a lotação do candidato
                                dsClassificados.Tables[0].Rows[i]["Lotação Final"] = (string)dr[0]["Unidade"];

                                //diminui uma vaga do local para onde ele foi
                                dr[0]["Vagas"] = Convert.ToInt32(dr[0]["Vagas"]) - 1;

                                //libera a vaga dele para o concurso de remoção
                                liberaVagaEmUnidade(tempLotacao);

                                //exclui as opções dele abaixo da atual
                                for (int k = j; k <= maxOpcoes; k++)
                                {
                                    if (dsOpcoes.Tables[0].Rows[i][k] != DBNull.Value && (string)dsOpcoes.Tables[0].Rows[i][k] != string.Empty)
                                    {
                                        dsOpcoes.Tables[0].Rows[i][k] = string.Empty;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                break;
                            }

                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }

            //percorrer lista de vagas abertas colocando-as novamente como opções no concurso de remoção
            foreach (String unid in lsVagasAbertas)
            {
                DataRow[] dr = dsVagas.Tables[0].Select("Unidade = '" + unid + "'");

                if (dr.Count() > 0)
                {
                    dr[0]["Vagas"] = Convert.ToInt32(dr[0]["Vagas"]) + 1;
                }
                else
                {
                    DataRow newDr = dsVagas.Tables[0].NewRow();
                    newDr["Unidade"] = unid;
                    newDr["Vagas"] = 1;
                    dsVagas.Tables[0].Rows.Add(newDr);
                }

            }
            lsVagasAbertas.Clear();
        }

        GridView1.DataSource = dsClassificados.Tables[0].DefaultView;
        GridView1.DataBind();
    }

    private void liberaVagaEmUnidade(String unidade)
    {
        String[] naoAbre = { "COGER", "ESCOR", "NUCOR", "COPEI", "ESPEI", "NUPEI", "DRJ" };
        String[] oc = { "0021902 - COTEC - COORDENAÇÃO-GERAL DE TECNOLOGIA DA INFORMAÇÃO",
                        "0021704 - COPES - COORDENAÇÃO-GERAL DE PROGRAMACAO E ESTUDOS",
                        "0021801 - COANA - COORDENAÇÃO-GERAL DE ADMINISTRAÇÃO ADUANEIRA",
                        "0010100 - GABINETE RFB",
                        "0010200 - ASSESSORIA ESPECIAL",
                        "0021700 - SUFIS - SUBSECRETARIA DE FISCALIZAÇÃO",
                        "0021701 - COFIS - COORDENAÇÃO-GERAL DE FISCALIZAÇÃO",
                        "0021600 - SUTRI - SUBSECRETARIA DE TRIBUTAÇÃO E CONTENCIOSO",
                        "0011000 - ASCOM - ASSESSORIA DE COMUNICAÇÃO SOCIAL",
                        "0010600 - COPAV - COORD. GERAL DE PLANEJ., ORGANIZAÇÃO E AVALIAÇÃO INSTI",
                        "0010800 - AUDIT - COORDENAÇÃO GERAL DE AUDITORIA INTERNA",
                        "0021500 - SUARA - SUBSECRETARIA DE ARRECADAÇÃO E ATENDIMENTO",
                        "0021501 - CODAC - COORDENAÇÃO-GERAL DE ARRECADAÇÃO E COBRANÇA",
                        "0021502 - COAEF - COORDENAÇÃO-GERAL DE ATENDIMENTO E EDUCAÇÃO FISCAL",
                        "0021503 - COCAD - COORDENAÇÃO-GERAL DE GESTÃO DE CADASTROS",
                        "0021504 - COREC - COORD. ESPECIAL DE RESSARCIMENTO, COMPENSAÇÃO E RESTIT",
                        "0021601 - COSIT - COORDENAÇÃO-GERAL DE TRIBUTAÇÃO",
                        "0021603 - COCAJ - COORDENAÇÃO-GERAL DE CONTENCIOSO ADMINISTRATIVO E JUDI",
                        "0021702 - COMAC - COORDENAÇÃO ESPECIAL DE MAIORES CONTRIBUINTES",
                        "0021800 - SUARI - SUBSECRETARIA DE ADUANA E RELAÇÕES INTERNACIONAIS",
                        "0021802 - CORIN - COORDENAÇÃO-GERAL DE RELAÇÕES INTERNACIONAIS",
                        "0011500 - COCIF/COORD-GERAL COOP E INTEGRACAO FISCAL",
                        "0021604 - COGET - COORD-GERAL DE ESTUDOS ECONÔMICO-TRIBUTÁRIOS E DE PREV. E ANÁLISE DE ARRECADAÇÃO",
                        "0021901 - COPOL - COORDENAÇÃO-GERAL DE PROGRAMAÇÃO E LOGÍSTICA",
                        "0021903 - COGEP - COORDENAÇÃO-GERAL DE GESTÃO DE PESSOAS",
                        "0021900 - SUCOR - SUBSECRETARIA DE GESTÃO CORPORATIVA"};
        foreach (String unid in naoAbre)
        {
            if (unidade.ToUpper().Contains(unid))
            {
                return;
            }
        }

        DataRow[] dr = dsEstudo.Tables[0].Select("UNIDADE = '" + unidade.Replace("'", "''") + "'");
        if (dr.Count() > 0)
        {
            if (!Convert.ToBoolean(dr[0]["ABREVAGA"]))
            {
                return;
            }
        }

        //Colocar na planilha: UNIDADE = UNIDADES CENTRAIS
        foreach (String unid in oc)
        {
            if (unidade.ToUpper().Equals(unid))
            {
                dr = dsEstudo.Tables[0].Select("UNIDADE = 'UNIDADES CENTRAIS'");
                if (dr.Count() > 0)
                {
                    if (!Convert.ToBoolean(dr[0]["ABREVAGA"]))
                    {
                        return;
                    }
                    else
                    {
                        break;
                    }
                }

            }
        }

        lsVagasAbertas.Add(unidade);
    }

    private String getStringDesformatada(String str)
    {
        String retorno = str.ToUpper().Replace("Ç", "C").Replace("Á", "A").Replace("Ã", "A").Replace("Ú", "U");
        retorno = retorno.Replace("É", "E").Replace("Ê", "E").Replace("Ó", "O").Replace("Õ", "O").Replace("Í", "I");
        retorno = retorno.Replace("'", "''");
        int x;
        if (Int32.TryParse(retorno.Substring(0, 7), out x))
        {
            retorno = retorno.Substring(10);
        }
        if (retorno.Contains("("))
        {
            retorno = retorno.Replace(retorno.Substring(retorno.IndexOf("(")), string.Empty);
        }
        return retorno;
    }

    //Ajustes a fazer na planilha Preliminar.xls:
    //Colocar Opção: 201, 201, ... 241, ...
    private void importExcelToGrid(bool auditor)
    {
        String strConn;
        String strConn3;
        string path = Server.MapPath("") + "\\planilhas\\";
        if (auditor)
        {
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "PreliminarAFRFB.xls;" +
                      "Extended Properties=Excel 8.0;";
            strConn3 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "DesistenciasAFRFB.xls;" +
                      "Extended Properties=Excel 8.0;";
            //strConn = "Provider=Microsoft.ACE.oledb.12.0;datasource=E:\\Elmo\\RFB\\Remocao\\2012\\PreliminarAFRFB.ods;";
        }
        else
        {
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "PreliminarATRFB.xls;" +
                      "Extended Properties=Excel 8.0;";
            strConn3 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "DesistenciasATRFB.xls;" +
                      "Extended Properties=Excel 8.0;";
        }

        String strConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;" +
        "Data Source=" + path + "GRAU DE LOTACAO AJUSTADO.xls;" +
        "Extended Properties=Excel 8.0;";

        OleDbDataAdapter daClass = new OleDbDataAdapter
        ("SELECT Clas, Nome, [Unidade de saída] as [Lotação Final], [Unidade de saída] as [Lotaçao Inicial] FROM [Classificacao$]", strConn);

        OleDbDataAdapter daDesistencias = new OleDbDataAdapter
        ("SELECT Class, [Nome do Candidato], CodUA FROM [Desisitencias]", strConn3);

        OleDbDataAdapter daOpcoes = new OleDbDataAdapter();
        daOpcoes.SelectCommand = new OleDbCommand();
        daOpcoes.SelectCommand.Connection = new OleDbConnection(strConn);
        daOpcoes.SelectCommand.CommandText = "SELECT Clas";
        for (int i = 1; i <= maxOpcoes; i++)
        {
            daOpcoes.SelectCommand.CommandText += ", [Opção: " + i.ToString() + "]";
        }
        daOpcoes.SelectCommand.CommandText += " FROM [Classificacao$]";

        OleDbDataAdapter daVagas = new OleDbDataAdapter
        ("SELECT Unidade, Vagas, [Unidade de Lotação], [Unidade de Exercício] FROM [Vagas$]", strConn2);

        OleDbDataAdapter daEstudo;

        if (auditor)
        {
            daEstudo = new OleDbDataAdapter
        ("SELECT UNIDADE, ABREVAGA_AF as ABREVAGA FROM [GL AJUSTADO$]", strConn2);
        }
        else
        {
            daEstudo = new OleDbDataAdapter
        ("SELECT UNIDADE, ABREVAGA_AT AS ABREVAGA FROM [GL AJUSTADO$]", strConn2);
        }

        daClass.Fill(dsClassificados);
        daOpcoes.Fill(dsOpcoes);
        daVagas.Fill(dsVagas);
        daEstudo.Fill(dsEstudo);
        try
        {
            daDesistencias.Fill(dsDesistencias);
        }
        catch
        {
            temDesistencias = false;
        }

        //foreach (DataRow dr in dsVagas.Tables[0].Rows)
        //{
        //    dr["Unidade de Lotação"] = getStringDesformatada(dr["Unidade de Lotação"].ToString());
        //    dr["Unidade de Exercício"] = getStringDesformatada(dr["Unidade de Exercício"].ToString());
        //}

        dsClassificados.Tables[0].Columns.Add("Removido");
        dsClassificados.Tables[0].Columns["Removido"].Expression = "IIF([Lotação Final] = [Lotaçao Inicial],'','Contemplado')";

        daClass.Dispose();
        daOpcoes.Dispose();
        daVagas.Dispose();
        daEstudo.Dispose();
        daDesistencias.Dispose();
    }

    private void inicioRecursaoPermutas()
    {

        while (rodarDeNovo)
        {
            rodarDeNovo = false;
            List<int> vetor1 = new List<int> { 0 };
            List<int> vetor2 = new List<int>();
            for (int i = 1; i < dsClassificados.Tables[0].Rows.Count; i++)
            {
                vetor2.Add(i);
            }
            permutacao(vetor1, vetor2, 0);
        }

    }

    private void permutacao(List<int> vetor1, List<int> vetor2, int indice)
    {
        for (int i = 0; i < vetor2.Count; i++)
        {
            for (int j = 1; j <= maxOpcoes; j++)
            {
                if (dsOpcoes.Tables[0].Rows[i]["Opção: " + j.ToString()] != DBNull.Value &&
                    !dsOpcoes.Tables[0].Rows[i]["Opção: " + j.ToString()].Equals(string.Empty))
                {
                    string local = (string)dsOpcoes.Tables[0].Rows[i]["Opção: " + j.ToString()];
                    if (local.Equals(dsClassificados.Tables[0].Rows[vetor1[indice]]["Lotação Final"]))
                    {
                        for (int k = 1; k <= maxOpcoes; k++)
                        {
                            if (dsOpcoes.Tables[0].Rows[vetor1[indice]]["Opção: " + k.ToString()] != DBNull.Value &&
                                !dsOpcoes.Tables[0].Rows[vetor1[indice]]["Opção: " + k.ToString()].Equals(string.Empty))
                            {
                                string local2 = (string)dsOpcoes.Tables[0].Rows[vetor1[indice]]["Opção: " + k.ToString()];
                                if (local2.Equals(dsClassificados.Tables[0].Rows[i]["Lotação Final"].ToString()))
                                {
                                    vetor1.Add(i);
                                    permutar(vetor1);
                                    rodarDeNovo = true;
                                    return;
                                }
                                else
                                {
                                    int[] vet3 = null;
                                    vetor1.CopyTo(vet3);
                                    List<int> vetor3 = new List<int>(vet3);
                                    int[] vet4 = null;
                                    vetor2.CopyTo(vet4);
                                    List<int> vetor4 = new List<int>(vet4);
                                    vetor3.Add(i);
                                    vetor4.Remove(i);
                                    permutacao(vetor3, vetor4, vetor1.Count);
                                    if (rodarDeNovo)
                                    {
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
                else
                {
                    break;
                }

            }
        }
    }

    private void permutar(List<int> vetor)
    {
        string lotacaoPrimeiro = (string)dsClassificados.Tables[0].Rows[vetor[0]]["Lotação Final"];
        for (int i=1; i<vetor.Count;i++)
        {
            dsClassificados.Tables[0].Rows[vetor[i - 1]]["Lotação Final"] = dsClassificados.Tables[0].Rows[vetor[i]]["Lotação Final"];

            for (int j = 1; j <= maxOpcoes; j++)
            {
                if (dsOpcoes.Tables[0].Rows[vetor[i - 1]]["Opção: " + j.ToString()] != DBNull.Value &&
                    !dsOpcoes.Tables[0].Rows[vetor[i - 1]]["Opção: " + j.ToString()].Equals(string.Empty))
                {
                    if (dsOpcoes.Tables[0].Rows[vetor[i - 1]]["Opção: " + j.ToString()].Equals(dsClassificados.Tables[0].Rows[vetor[i]]["Lotação Final"].ToString()))
                    {
                        for (int k = j; k <= maxOpcoes; k++)
                        {
                            if (dsOpcoes.Tables[0].Rows[vetor[i - 1]][k] != DBNull.Value && (string)dsOpcoes.Tables[0].Rows[vetor[i - 1]][k] != string.Empty)
                            {
                                dsOpcoes.Tables[0].Rows[vetor[i - 1]][k] = string.Empty;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
                else
                {
                    break;
                }
            }
        }
        dsClassificados.Tables[0].Rows[vetor[vetor.Count-1]]["Lotação Final"] = lotacaoPrimeiro;
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {
        string cargo = "";
        if (rbFiscal.Checked)
        {
            cargo = "AFRFB";
        }
        else
        {
            cargo = "ATRFB";
        }

        string path = Server.MapPath("") + "\\planilhas\\";
        if (fuEglVagas.HasFile)
        {

                string fileExtension = System.IO.Path.GetExtension(fuEglVagas.PostedFile.FileName).ToLower();
                if (fileExtension == ".xls")
                {
                    fuEglVagas.PostedFile.SaveAs(path + "GRAU DE LOTACAO AJUSTADO.xls");
                }
        }
        if (fuClasPreliminar.HasFile)
        {

            string fileExtension = System.IO.Path.GetExtension(fuClasPreliminar.PostedFile.FileName).ToLower();
            if (fileExtension == ".xls")
            {
                fuClasPreliminar.PostedFile.SaveAs(path + "Preliminar" + cargo + ".xls");
            }
        }
        if (fuDesist.HasFile)
        {

            string fileExtension = System.IO.Path.GetExtension(fuDesist.PostedFile.FileName).ToLower();
            if (fileExtension == ".xls")
            {
                fuDesist.PostedFile.SaveAs(path + "Desistencias" + cargo + ".xls");
            }
        }
    }
    protected void btnSimular_Click(object sender, EventArgs e)
    {
        if (rbFiscal.Checked)
        {
            importExcelToGrid(true);
        }
        else
        {
            importExcelToGrid(false);
        }
        simular();
        //inicioRecursaoPermutas();
    }
}