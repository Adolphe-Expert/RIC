Option Explicit On
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.OleDb.OleDbConnection
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

'1111111111111111111111111111111111111
'1111111111111111111111111111111111111

'Imports excel = Microsoft.Office.Interop.Excel

'Public Class ClasseExcel
'    Private objexcel As New excel.Application
'    Dim xlBook As excel.Workbook
'    Dim xlworksheet As excel.Worksheet
'    Public Sub New()
'        xlBook = objexcel.Workbooks.Add
'        xlworksheet = CType(xlBook.ActiveSheet, excel.Worksheet)
'    End Sub

'    Public Sub Writelist(ByVal mylist As List(Of String))
'        ' écrit dans la colonne A1 :: A?  mylist
'        Dim cellstrcopy As String
'        Dim indexcol As Integer
'        Dim indexrow As Integer
'        Dim myfont As Font
'        myfont = New Font("arial", 12, FontStyle.Bold)
'        cellstrcopy = String.Empty
'        Try
'            With xlworksheet
'                indexcol = 1
'                indexrow = 1
'                cellstrcopy = Convert.ToChar(indexcol + 64) & indexrow.ToString
'                For Each item In mylist

'                    .Cells(indexrow, indexcol) = item

'                    .Range(cellstrcopy).BorderAround()
'                    .Range(cellstrcopy).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon)
'                    .Range(cellstrcopy).Select()
'                    .Range(cellstrcopy).HorizontalAlignment = excel.XlVAlign.xlVAlignCenter
'                    With .Range(cellstrcopy).Font
'                        .Name = "Arial"
'                        .Strikethrough = False
'                        .Bold = True
'                        .Size = 12
'                    End With
'                    indexrow += 1
'                    cellstrcopy = Convert.ToChar(indexcol + 64) & indexrow.ToString
'                Next

'            End With
'            objexcel.Visible = True
'            objexcel = Nothing
'        Catch ex As Exception
'            MessageBox.Show(ex.Message.ToString)
'        End Try

'    End Sub
'End Class

'1111111111111111111111111111111111111
'1111111111111111111111111111111111111

Public Class Form1

    Dim xlAppSource As Excel.Application
    Dim xlWorkBookSource As Excel.Workbook
    Dim xlWorkSheetSource As Excel.Worksheet

    Dim xlAppRapport As Excel.Application
    Dim xlWorkBookRapport As Excel.Workbook
    Dim xlWorkSheetRapport As Excel.Worksheet

    Dim RepportPath As String

    Dim Affichage As String
    Dim NombreSup As Integer
    'Affichage = "Null"

    'Dim Truc As CheckBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'yes yes
        'Dim TreeViewFiltresGeographiques As TreeView
        'TreeViewFiltresGeographiques = New TreeView()
        'Me.Controls.Add(TreeViewFiltresGeographiques)
        'TreeViewFiltresGeographiques.Nodes.Clear()
        'TreeViewFiltresGeographiques.SelectedNode.Checked = True

        'For Each Truc In GroupBoxChoixInfras.Controls
        '    If Truc.Checked = True Then
        '        ComboBoxInfraAFiltrer.Text = Truc.Text
        '    End If
        'Next Truc

        GroupBoxAdministration.Hide()
        GroupBoxEauEtAssainissement.Hide()
        GroupBoxSante.Hide()
        GroupBoxEducation.Hide()
        GroupBoxAgricultureEtElevage.Hide()
        GroupBoxAgricultureEtElevage.Hide()
        GroupBoxTransport.Hide()
        GroupBoxEquipementsMarchands.Hide()
        GroupBoxEnergie.Hide()
        GroupBoxSportsEtLoisirs.Hide()

        'GroupBoxFiltresGeographiques.Hide()
        'GroupBoxFiltresSpecifiquesToutInfras.Hide()

        'Dim myexcel As New ClasseExcel
        'Dim lalist As New List(Of String)
        'For iter = 0 To 20
        '    lalist.Add("toto" & iter.ToString)
        'Next
        'myexcel.Writelist(lalist)

        GenererRapportExcel(sFile:=" D:\Education v1.xlsx")

    End Sub

    Private Sub ButtonSelectionnerToutInfra_Click(sender As Object, e As EventArgs) Handles ButtonSelectionnerToutInfra.Click
        CheckBoxAdministration.Checked = True
        CheckBoxAgriculture.Checked = True
        CheckBoxEauAssainissement.Checked = True
        CheckBoxEducation.Checked = True
        CheckBoxElevage.Checked = True
        CheckBoxEnergie.Checked = True
        CheckBoxEquipementMarchand.Checked = True
        CheckBoxSante.Checked = True
        CheckBoxSportsLoisirs.Checked = True
        CheckBoxTransport.Checked = True

        'For Each Truc In GroupBox1.Controls
        '    Truc.Checked = True
        'Next Truc


    End Sub

    Private Sub ButtonReinitialiserInfra_Click(sender As Object, e As EventArgs) Handles ButtonReinitialiserInfra.Click
        CheckBoxAdministration.Checked = False
        CheckBoxAgriculture.Checked = False
        CheckBoxEauAssainissement.Checked = False
        CheckBoxEducation.Checked = False
        CheckBoxElevage.Checked = False
        CheckBoxEnergie.Checked = False
        CheckBoxEquipementMarchand.Checked = False
        CheckBoxSante.Checked = False
        CheckBoxSportsLoisirs.Checked = False
        CheckBoxTransport.Checked = False

    End Sub

    'Private Sub TreeViewFiltresGeographiques_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterSelect
    '    If TreeViewFiltresGeographiques.SelectedNode.Checked = False Then
    '        TreeViewFiltresGeographiques.SelectedNode.Checked = True
    '    ElseIf TreeViewFiltresGeographiques.SelectedNode.Checked = True Then
    '        TreeViewFiltresGeographiques.SelectedNode.Checked = False
    '    End If
    'End Sub


    '***************************************

    ' Updates all child tree nodes recursively.
    Private Sub CheckAllChildNodes(treeNode As TreeNode, nodeChecked As Boolean)
        Dim node As TreeNode
        For Each node In treeNode.Nodes
            node.Checked = nodeChecked
            If node.Nodes.Count > 0 Then
                ' If the current node has child nodes, call the CheckAllChildsNodes method recursively.
                Me.CheckAllChildNodes(node, nodeChecked)
            End If
        Next node
    End Sub

    ' NOTE   This code can be added to the BeforeCheck event handler instead of the AfterCheck event.
    ' After a tree node's Checked property is changed, all its child nodes are updated to the same value.
    Private Sub node_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterCheck
        ' The code only executes if the user caused the checked state to change.
        If e.Action <> TreeViewAction.Unknown Then
            If e.Node.Nodes.Count > 0 Then
                ' Calls the CheckAllChildNodes method, passing in the current 
                ' Checked value of the TreeNode whose checked state changed. 
                Me.CheckAllChildNodes(e.Node, e.Node.Checked)
            End If
        End If
    End Sub

    '///////////////////////////////////////

    Private Sub ComboBoxInfraAFiltrer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxInfraAFiltrer.SelectedIndexChanged

        If ComboBoxInfraAFiltrer.Text = "Eau et Assainissement" Then
            Me.Controls.Add(GroupBoxEauEtAssainissement)
            GroupBoxEauEtAssainissement.Show()
            GroupBoxEauEtAssainissement.BringToFront()
            GroupBoxEauEtAssainissement.Visible = True
            Affichage = "Eau et Assainissement"
            '------------------------------------
            GroupBoxSportsEtLoisirs.Hide()
            GroupBoxAdministration.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Sports et loisirs" Then
            Me.Controls.Add(GroupBoxSportsEtLoisirs)
            GroupBoxSportsEtLoisirs.Show()
            GroupBoxSportsEtLoisirs.BringToFront()
            GroupBoxSportsEtLoisirs.Visible = True
            Affichage = "Sports et loisirs"
            '------------------------------------
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxAdministration.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Administration" Then
            Me.Controls.Add(GroupBoxAdministration)
            GroupBoxAdministration.Show()
            GroupBoxAdministration.BringToFront()
            Affichage = "Administration"
            '------------------------------------
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()

            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Santé" Then
            Me.Controls.Add(GroupBoxSante)
            GroupBoxSante.Show()
            GroupBoxSante.BringToFront()
            Affichage = "Santé"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            'GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Education" Then
            Me.Controls.Add(GroupBoxEducation)
            GroupBoxEducation.Show()
            GroupBoxEducation.BringToFront()
            Affichage = "Education"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            'GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Agriculture" Then
            Me.Controls.Add(GroupBoxAgricultureEtElevage)
            GroupBoxAgricultureEtElevage.Show()
            GroupBoxAgricultureEtElevage.BringToFront()
            Affichage = "Agriculture"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Elevage" Then
            Me.Controls.Add(GroupBoxAgricultureEtElevage)
            GroupBoxAgricultureEtElevage.Show()
            GroupBoxAgricultureEtElevage.BringToFront()
            Affichage = "Elevage"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Transport" Then
            Me.Controls.Add(GroupBoxTransport)
            GroupBoxTransport.Show()
            GroupBoxTransport.BringToFront()
            Affichage = "Transport"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Equipements marchands" Then
            Me.Controls.Add(GroupBoxEquipementsMarchands)
            GroupBoxEquipementsMarchands.Show()
            GroupBoxEquipementsMarchands.BringToFront()
            Affichage = "Equipements marchands"

            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            'GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Energie" Then
            Me.Controls.Add(GroupBoxEnergie)
            GroupBoxEnergie.Show()
            GroupBoxEnergie.BringToFront()
            Affichage = "Energie"

            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            'GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        End If
    End Sub
    '////////////////////////////////////////////

    Private Sub ButtonSelectAllFiltres_Click(sender As Object, e As EventArgs) Handles ButtonSelectAllFiltres.Click
        If Affichage = "Education" Then
            CheckBoxEcolePrive.Checked = True
            CheckBoxEcolePublique.Checked = True
            CheckBoxBonEtatEducation.Checked = True
            CheckBoxEtatUsageEducation.Checked = True
            CheckBoxInutilisableEducation.Checked = True
            CheckBoxPrescolaire.Checked = True
            CheckBoxPrimaire.Checked = True
            CheckBoxCollege.Checked = True
            CheckBoxLycee.Checked = True
            NombreSup = NumericUpDown1.Value


        End If
    End Sub

    Private Sub ButtonCancelFiltres_Click(sender As Object, e As EventArgs) Handles ButtonCancelFiltres.Click
        If Affichage = "Education" Then
            CheckBoxEcolePrive.Checked = False
            CheckBoxEcolePublique.Checked = False
            CheckBoxBonEtatEducation.Checked = False
            CheckBoxEtatUsageEducation.Checked = False
            CheckBoxInutilisableEducation.Checked = False
            CheckBoxPrescolaire.Checked = False
            CheckBoxPrimaire.Checked = False
            CheckBoxCollege.Checked = False
            CheckBoxLycee.Checked = False
            NombreSup = 0
            NumericUpDown1.Value = 0

        End If
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Private Sub ButtonVoirLesFiltres_Click(sender As Object, e As EventArgs) Handles ButtonVoirLesFiltres.Click
    '    'GroupBoxFiltresGeographiques.Show()
    '    'GroupBoxFiltresSpecifiquesToutInfras.Show()
    '    ComboBoxInfraAFiltrer.Items.Clear()
    '    ComboBoxInfraAFiltrer.Text = ""
    '    If CheckBoxAdministration.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Administration")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Administration")
    '    End If
    '    If CheckBoxEducation.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Education")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Education")
    '    End If
    '    If CheckBoxSportsLoisirs.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Sports et loisirs")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Sports et loisirs")
    '    End If
    '    If CheckBoxEauAssainissement.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Eau et Assainissement")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Eau et Assainissement")
    '    End If
    '    If CheckBoxSante.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Santé")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Santé")
    '    End If
    '    If CheckBoxAgriculture.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Agriculture")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Agriculture")
    '    End If
    '    If CheckBoxElevage.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Elevage")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Elevage")
    '    End If
    '    If CheckBoxTransport.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Transport")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Transport")
    '    End If
    '    If CheckBoxEquipementMarchand.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Equipements marchands")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Equipements marchands")
    '    End If
    '    If CheckBoxEnergie.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Energie")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Energie")
    '    End If
    'End Sub

    Private Sub CheckBoxAdministration_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAdministration.CheckedChanged
        'ComboBoxInfraAFiltrer.Items.Clear()
        'ComboBoxInfraAFiltrer.Text = ""
        If CheckBoxAdministration.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Administration")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Administration")
        End If
    End Sub

    Private Sub CheckBoxAgriculture_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAgriculture.CheckedChanged
        If CheckBoxAgriculture.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Agriculture")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Agriculture")
        End If
    End Sub

    Private Sub CheckBoxEauAssainissement_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEauAssainissement.CheckedChanged
        If CheckBoxEauAssainissement.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Eau et Assainissement")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Eau et Assainissement")
        End If
    End Sub

    Private Sub CheckBoxElevage_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxElevage.CheckedChanged
        If CheckBoxElevage.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Elevage")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Elevage")
        End If
    End Sub

    Private Sub CheckBoxEducation_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEducation.CheckedChanged
        If CheckBoxEducation.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Education")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Education")
        End If
    End Sub

    Private Sub CheckBoxEquipementMarchand_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEquipementMarchand.CheckedChanged
        If CheckBoxEquipementMarchand.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Equipements marchands")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Equipements marchands")
        End If
    End Sub

    Private Sub CheckBoxEnergie_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEnergie.CheckedChanged
        If CheckBoxEnergie.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Energie")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Energie")
        End If
    End Sub

    Private Sub CheckBoxSante_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSante.CheckedChanged
        If CheckBoxSante.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Santé")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Santé")
        End If

    End Sub

    Private Sub CheckBoxSportsLoisirs_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSportsLoisirs.CheckedChanged
        If CheckBoxSportsLoisirs.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Sports et loisirs")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Sports et loisirs")
        End If
    End Sub

    Private Sub CheckBoxTransport_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxTransport.CheckedChanged
        If CheckBoxTransport.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Transport")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Transport")

            'ComboBoxInfraAFiltrer.Items.Add("Transport")
            'ComboBoxInfraAFiltrer.Enabled = False
            'ComboBoxInfraAFiltrer.SelectedItem = False

        End If
    End Sub

    Private Sub TreeViewFiltresGeographiques_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterSelect

    End Sub

    Private Sub EducationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EducationToolStripMenuItem.Click
        ' OpenFileDialogEducation.ShowDialog()
        With OpenFileDialogEducation
            .Title = "Fichier excel source pour le RIC - Infrastructure ''EDUCATION'' "           ' DIALOG BOX TITLE.
            .FileName = ""
            .Filter = "Fichier Excel du RIC|*.xlsx;*.xls"     ' FILTER ONLY EXCEL FILES IN FILE TYPE.
            Dim sFileName As String = ""

            If .ShowDialog() = DialogResult.OK Then
                sFileName = .FileName

                'If Trim(sFileName) <> "" Then
                '    EditEmpDetails(sFileName)       ' PROCEDURE TO EDIT EMPLOYEE DETAILS.
                'End If
                LabelSourceEducation.Text = OpenFileDialogEducation.FileName
            End If

        End With
    End Sub

    'Private Sub LoadData()
    '    con = New OleDbConnection(conString)
    '    Dim query As String = "SELECT * FROM [EDUCATION$] "
    '    adapter = New OleDbDataAdapter(query, con)

    '    Dim ds As DataSet = New DataSet()
    '    Dim dt As DataTable = New DataTable

    '    adapter.Fill(dt)

    '    DataGridView1.DataSource = ds.Tables(0)
    '    DataGridView1.DataMember = "[EDUCATION$]"

    '    con.Close()

    '    '------------------------------------------

    '    ''déclaration du dataset
    '    'Dim dat As DataSet
    '    'dat = New DataSet
    '    ''déclaration et utilisation d'un OLeDBConnection
    '    'Using Conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\EDUCATION - labels - 2017-11-22-14-22.xlsx';Extended Properties=""Excel 8.0;HDR=Yes;""")
    '    '    ' Conn.Open()
    '    '    'déclaration du DataAdapter
    '    '    'notre requête sélectionne toute les cellule de la Feuil1
    '    '    Using Adap As OleDbDataAdapter = New OleDbDataAdapter("select * from [EDUCATION$]", Conn)
    '    '        'Chargement du Dataset
    '    '        Adap.Fill(dat)
    '    '        'On Binde les données sur le DGV
    '    '        DataGridView1.DataSource = dat.Tables(0)
    '    '    End Using
    '    '    'le end using libère les ressources
    '    'End Using
    '    '------------------------------------------

    'End Sub

    Private Sub ButtonAppliquerFiltres_Click(sender As Object, e As EventArgs) Handles ButtonAppliquerFiltres.Click

        '1st method start /////////////////////////////////////////////////
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

            'Dim path As String = "D:\EDUCATION - labels - 2017-11-22-14-22.xlsx"

            Dim path As String = LabelSourceEducation.Text
            'LabelSourceEducation.Text definit la source des données

            'Dim sourceDesDonnees As String = path

            'If LabelSourceEducation.Text = "" Then

            'End If


            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;""")
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [EDUCATION$] Where Eau_potable_accesssible_sur_le = 'Oui' ", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [EDUCATION$]", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [DISTRICT], Nom_de_l_tablissement_scolaire, Type_d_tablissement, Niveaux_enseign_s from [EDUCATION$]", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [_index], [DISTRICT], [Nom_de_l_tablissement_scolaire], [Type_d_tablissement], [Niveaux_enseign_s] from [EDUCATION$]", MyConnection)
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select Reference_enquete,DISTRICT, Nom_de_l_tablissement_scolaire, Type_d_tablissement, Niveaux_enseign_s from [EDUCATION$]", MyConnection)

            DataSet = New System.Data.DataSet

            MyCommand.Fill(DataSet)

            DataGridView1.DataSource = DataSet.Tables(0)

            Dim chk As New DataGridViewCheckBoxColumn()
            'Dim gridButtom As New DataGridViewLinkColumn

            DataGridView1.Columns.Add(chk)

            chk.HeaderText = "Check Data"

            chk.Name = "chk"

            Dim gridButtom As New DataGridViewButtonColumn

            DataGridView1.Columns.Add(gridButtom)

            gridButtom.HeaderText = "gridButtom"

            gridButtom.Name = "gridButtom"

            'DataGridView1.Rows(2).Cells(3).Value = True
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        '1st method end ////////////////////////////////////////////////////

        'LoadData()

        '2nd method start ////////////////////////////////////////////////////


        'Dim con As New OleDbConnection
        'Dim cm As New OleDbCommand
        'con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\crysol\Desktop\TEST\Book1.xls;Extended Properties=""Excel 12.0 Xml;HDR=YES""")
        'con.Open()
        'With cm
        '    .Connection = con
        '    .CommandText = "update [up$] set [name]=?, [QC_status]=?, [reason]=?, [date]=? WHERE [article_no]=?"
        '    cm = New OleDbCommand(.CommandText, con)
        '    cm.Parameters.AddWithValue("?", TextBox2.Text)
        '    cm.Parameters.AddWithValue("?", ComboBox1.SelectedItem)
        '    cm.Parameters.AddWithValue("?", TextBox3.Text)
        '    cm.Parameters.AddWithValue("?", DateTimePicker1.Text)
        '    cm.Parameters.AddWithValue("?", TextBox1.Text)
        '    cm.ExecuteNonQuery()
        '    MsgBox("UPDATE SUCCESSFUL")
        '    con.Close()
        'End With

        '2nd method end  /////////////////////////////////////////////////////

    End Sub

    Private Sub ButtonVoirRapportDetaille_Click(sender As Object, e As EventArgs) Handles ButtonVoirRapportDetaille.Click

        '///////////////////'1st method start

        'Try
        '    Dim oXL As Excel.Application
        '    Dim oWB As Excel.Workbook
        '    Dim oSheet As Excel.Worksheet
        '    Dim oRng As Excel.Range
        '    'On Error GoTo Err_Handler
        '    ' Start Excel and get Application object.
        '    oXL = New Excel.Application

        '    'Get a new workbook.
        '    Dim path As String = ViewState("filepath")
        '    oWB = oXL.Workbooks.Open(path)
        '    oSheet = CType(oXL.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value), Excel.Worksheet)
        '    'oSheet.Name = "Reject_History"
        '    Dim totalSheets As Integer = oXL.ActiveWorkbook.Sheets.Count
        '    CType(oXL.ActiveSheet, Excel.Worksheet).Move(After:=oXL.Worksheets(totalSheets))
        '    CType(oXL.ActiveWorkbook.Sheets(totalSheets), Excel.Worksheet).Activate()

        '    'Write Dataset to Excel Sheet
        '    Dim col As Integer = 0

        '    For Each dr As DataColumn In DirectCast(ViewState("DisplayNonExisting"), DataTable).Columns

        '        col += 1
        '        'Determine cell to write
        '        oSheet.Cells(10, col).Value = dr.ColumnName

        '    Next

        '    Dim irow As Integer = 10
        '    For Each dr As DataRow In DirectCast(ViewState("DisplayNonExisting"), DataTable).Rows
        '        irow += 1
        '        Dim icol As Integer = 0
        '        For Each c As String In dr.ItemArray
        '            icol += 1
        '            'Determine cell to write
        '            oSheet.Cells(irow, icol).Value = c
        '        Next
        '    Next

        '    ' Make sure Excel is visible and give the user control
        '    ' of Microsoft Excel's lifetime.
        '    ' oXL.Visible = True
        '    ' oXL.UserControl = True

        '    'oWB.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
        '    'oWB.Close()
        '    oWB.Save()
        '    oWB.Close(Type.Missing, Type.Missing, Type.Missing)
        '    ' Make sure you release object references.

        '    oRng = Nothing
        '    oSheet = Nothing
        '    oWB = Nothing
        '    oXL = Nothing

        'Catch ex As Exception

        'End Try
        '///////////////////'1st method end

        '*************************** 2nd method

        '*************************** 2nd method
    End Sub

    ' EDIT DETAILS IN THE EXCEL FILE.
    Private Sub GenererRapportExcel(ByVal sFile As String)
        ' THE EXCEL NAMESPACE ALLOWS US TO USE THE EXCEL APPLICATION CLASS

        xlAppSource = New Excel.Application
        xlWorkBookSource = xlAppSource.Workbooks.Open(sFile)           ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheetSource = xlWorkBookSource.Worksheets("EDUCATION")    ' THE NAME OF THE WORK SHEET. 

        'sFile = "D:\Education v1.xlsx"

        Dim iRow As Integer = 0
        Dim iCol As Integer = 0

        For iRow = 2 To xlWorkSheetSource.Rows.Count
            If Trim(xlWorkSheetSource.Cells(iRow, 1).value) = "" Then
                Exit For        ' BAIL OUT IF REACHED THE LAST ROW.
            End If

            For iCol = 1 To xlWorkSheetSource.Columns.Count
                If Trim(xlWorkSheetSource.Cells(1, iCol).value) = "" Then
                    Exit For    ' BAIL OUT IF REACHED THE LAST COLUMN.
                End If

                ' CHECK IF THE SELECTED EMPLOYEE EXISTS AND CHANGE THE MOBILE NO.
                If Trim(xlWorkSheetSource.Cells(iRow, iCol).value) = Trim("cmbEmp.Text") Then
                    xlWorkSheetSource.Cells(iRow, iCol + 1) = Trim("tbMobile.Text")
                    Exit For    ' DONE. GET OUT OF THE LOOP.
                End If
            Next
        Next

        xlWorkBookSource.Close() : xlAppSource.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSource) : xlAppSource = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBookSource) : xlWorkBookSource = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheetSource) : xlWorkSheetSource = Nothing
    End Sub
    'End Class

    '////////////////////////////////////////////

End Class
