<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server" enctype="multipart/form-data">
        <asp:RadioButton ID="rbFiscal" runat="server" Checked="True" GroupName="Cargo" Text="Fiscal" />
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:RadioButton ID="rbAnalista" runat="server" GroupName="Cargo" Text="Analista" />
        <br />
        Selecione planilhas para fazer upload, se necessário:<br />
        Estudo de Lotação e Vagas:
        <asp:FileUpload ID="fuEglVagas" runat="server" />
        <br />
        Classificação Preliminar:
        <asp:FileUpload ID="fuClasPreliminar" runat="server" />
        <br />
        Desistências:
        <asp:FileUpload ID="fuDesist" runat="server" />
        <br />
        <br />
        Planilhas atuais:<br />
        <a href="planilhas/GRAU DE LOTACAO AJUSTADO.xls">Estudo de Lotação e Vagas</a>
        <br />
        <a href="planilhas/PreliminarAFRFB.xls">Classificação Preliminar</a>
        <br />
        <a href="planilhas/DesistenciasAFRFB.xls">Desistências</a>
        <br />
        <br />
        <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Upload" />
        <br />
        <br />
        <asp:Button ID="btnSimular" runat="server" Text="Simular"
            OnClick="btnSimular_Click" />
        <div>
            <asp:GridView ID="GridView1" runat="server">
            </asp:GridView>
        </div>
    </form>
</body>
</html>
