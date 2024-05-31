Attribute VB_Name = "Navegacao"
Sub PainelProd()
    If Sheets("PAINEL.PROD").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("PAINEL.PROD").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("PAINEL.PROD").Select
    Range("d3").Select
End Sub


Sub BoletimDiario()
    If Sheets("Boletim Diario").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Boletim Diario").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Boletim Diario").Select
    Range("a1").Select
End Sub



Sub OrcamentoProd()
    If Sheets("Ppto Mes").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Ppto Mes").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Ppto Mes").Select
    Range("b5").Select
End Sub

Sub MetaProd()
    If Sheets("Metas").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Metas").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Metas").Select
    Range("D2").Select
End Sub

Sub ProgDiario()
    If Sheets("Programa").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Programa").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Programa").Select
    Range("D2").Select
End Sub


Sub BIDiario()
    If Sheets("B.Diario").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Diario").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Diario").Select
    Range("B2").Select
    Selection.End(xlToRight).Select
End Sub

Sub BISemanal()
    If Sheets("B.Semanal").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Semanal").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Semanal").Select
    Range("B2").Select
    Selection.End(xlToRight).Select
End Sub

Sub BIMensal()
    If Sheets("B.Mensal").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Mensal").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Mensal").Select
    Range("B2").Select
    Selection.End(xlToRight).Select
End Sub

Sub BIAcum()
    If Sheets("B.Acum").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Acum").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Acum").Select
    Range("B2").Select
    Selection.End(xlToRight).Select
End Sub



Sub BPlantio()
    If Sheets("B.Campo").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Campo").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Campo").Select
    Range("v4").Select
     Selection.End(xlDown).Select
End Sub
Sub BChuva()
    If Sheets("B.Campo").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("B.Campo").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("B.Campo").Select
    Range("Ak4").Select
     Selection.End(xlDown).Select
End Sub

Sub BAnidro()
    If Sheets("Anidro").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Anidro").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Anidro").Select
    Range("H5").Select
    Selection.End(xlDown).Select
End Sub

Sub Bhidratado()
    If Sheets("Hidratado").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Hidratado").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Hidratado").Select
    Range("H5").Select
    Selection.End(xlDown).Select
End Sub

Sub BBagaco()
    If Sheets("Bagaço").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Bagaço").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Bagaço").Select
    Range("H5").Select
    Selection.End(xlDown).Select
End Sub

Sub CEPEA()
    If Sheets("CEPEA").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("CEPEA").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("CEPEA").Select
    Range("B4").Select
    Selection.End(xlDown).Select
End Sub

Sub Iventario()
    If Sheets("Inventario").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Inventario").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Inventario").Select
    Range("D6").Select
    Selection.End(xlDown).Select
End Sub

Sub Seguranca()
    If Sheets("Segurança").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Segurança").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Segurança").Select
    Range("h3").Select
    Selection.End(xlDown).Select
End Sub

Sub Paradas()
    If Sheets("Paradas").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Paradas").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Paradas").Select
    Range("A4").Select
    Selection.End(xlDown).Select
End Sub


Sub PainelMoagem()
    If Sheets("Painel Moagem").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Painel Moagem").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Painel Moagem").Select
    Range("d4").Select

End Sub

Sub PainelParadas()
    If Sheets("Painel Paradas").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("Painel Paradas").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("Painel Paradas").Select
    Range("A6").Select

End Sub
Sub IndAgricola()
    If Sheets("IndAgricola").Visible = xlSheetVeryHidden Then 'Se não estiver Oculta
        Sheets("IndAgricola").Visible = xlSheetVisible ' Então exibir
    End If
    Sheets("IndAgricola").Select
    Range("B3").Select

End Sub


Sub Senha()
FormSenha.Show
End Sub





