<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.TreeViewFiltresGeographiques = New System.Windows.Forms.TreeView()
        Me.GroupBoxFiltresGeographiques = New System.Windows.Forms.GroupBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImporterSourceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EauEtAssainissementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SportsEtLoisirsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdministrationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SantéToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EducationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AgricultureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ElevageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EquipementsMarchandsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EnergieToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AProposDeLoutilToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuitterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBoxSelectFiltres = New System.Windows.Forms.GroupBox()
        Me.ButtonCancelFiltres = New System.Windows.Forms.Button()
        Me.ButtonSelectAllFiltres = New System.Windows.Forms.Button()
        Me.ButtonAppliquerFiltres = New System.Windows.Forms.Button()
        Me.ComboBoxInfraAFiltrer = New System.Windows.Forms.ComboBox()
        Me.GroupBoxEducation = New System.Windows.Forms.GroupBox()
        Me.GroupBoxNombreDEleve = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.LabelNombreEleveSuperieurA = New System.Windows.Forms.Label()
        Me.GroupBoxNiveauEnseigne = New System.Windows.Forms.GroupBox()
        Me.CheckBoxLycee = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCollege = New System.Windows.Forms.CheckBox()
        Me.CheckBoxPrimaire = New System.Windows.Forms.CheckBox()
        Me.CheckBoxPrescolaire = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxEcolePublique = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEcolePrive = New System.Windows.Forms.CheckBox()
        Me.GroupBoxEtatInfraEducation = New System.Windows.Forms.GroupBox()
        Me.CheckBoxInutilisableEducation = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEtatUsageEducation = New System.Windows.Forms.CheckBox()
        Me.CheckBoxBonEtatEducation = New System.Windows.Forms.CheckBox()
        Me.GroupBoxAdministration = New System.Windows.Forms.GroupBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.CheckBox8 = New System.Windows.Forms.CheckBox()
        Me.CheckBox9 = New System.Windows.Forms.CheckBox()
        Me.CheckBox10 = New System.Windows.Forms.CheckBox()
        Me.CheckBox14 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxSante = New System.Windows.Forms.GroupBox()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown3 = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBoxDiplomeDuReferent = New System.Windows.Forms.GroupBox()
        Me.CheckBoxInfirmier = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSageFemme = New System.Windows.Forms.CheckBox()
        Me.CheckBoxDocteur = New System.Windows.Forms.CheckBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxCSBII = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCSBI = New System.Windows.Forms.CheckBox()
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.CheckBox19 = New System.Windows.Forms.CheckBox()
        Me.CheckBox20 = New System.Windows.Forms.CheckBox()
        Me.GroupBox14 = New System.Windows.Forms.GroupBox()
        Me.CheckBox21 = New System.Windows.Forms.CheckBox()
        Me.CheckBox22 = New System.Windows.Forms.CheckBox()
        Me.CheckBox23 = New System.Windows.Forms.CheckBox()
        Me.CheckBox24 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxEauEtAssainissement = New System.Windows.Forms.GroupBox()
        Me.GroupBox15 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown4 = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBoxTypeDInstallation = New System.Windows.Forms.GroupBox()
        Me.CheckBoxChâteauDEau = New System.Windows.Forms.CheckBox()
        Me.CheckBoxPuits = New System.Windows.Forms.CheckBox()
        Me.CheckBoxBorneFontaine = New System.Windows.Forms.CheckBox()
        Me.GroupBox17 = New System.Windows.Forms.GroupBox()
        Me.CheckBox18 = New System.Windows.Forms.CheckBox()
        Me.CheckBox25 = New System.Windows.Forms.CheckBox()
        Me.GroupBox18 = New System.Windows.Forms.GroupBox()
        Me.CheckBox26 = New System.Windows.Forms.CheckBox()
        Me.CheckBox27 = New System.Windows.Forms.CheckBox()
        Me.GroupBox19 = New System.Windows.Forms.GroupBox()
        Me.CheckBox28 = New System.Windows.Forms.CheckBox()
        Me.CheckBox29 = New System.Windows.Forms.CheckBox()
        Me.CheckBox30 = New System.Windows.Forms.CheckBox()
        Me.CheckBox31 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxSportsEtLoisirs = New System.Windows.Forms.GroupBox()
        Me.GroupBox16 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown5 = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox20 = New System.Windows.Forms.GroupBox()
        Me.CheckBox15 = New System.Windows.Forms.CheckBox()
        Me.CheckBox16 = New System.Windows.Forms.CheckBox()
        Me.CheckBox17 = New System.Windows.Forms.CheckBox()
        Me.GroupBox21 = New System.Windows.Forms.GroupBox()
        Me.CheckBox32 = New System.Windows.Forms.CheckBox()
        Me.CheckBox33 = New System.Windows.Forms.CheckBox()
        Me.GroupBox22 = New System.Windows.Forms.GroupBox()
        Me.CheckBox34 = New System.Windows.Forms.CheckBox()
        Me.CheckBox35 = New System.Windows.Forms.CheckBox()
        Me.GroupBox23 = New System.Windows.Forms.GroupBox()
        Me.CheckBox36 = New System.Windows.Forms.CheckBox()
        Me.CheckBox37 = New System.Windows.Forms.CheckBox()
        Me.CheckBox38 = New System.Windows.Forms.CheckBox()
        Me.CheckBox39 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxAgricultureEtElevage = New System.Windows.Forms.GroupBox()
        Me.GroupBox24 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown6 = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox25 = New System.Windows.Forms.GroupBox()
        Me.CheckBox40 = New System.Windows.Forms.CheckBox()
        Me.CheckBox41 = New System.Windows.Forms.CheckBox()
        Me.CheckBox42 = New System.Windows.Forms.CheckBox()
        Me.GroupBox26 = New System.Windows.Forms.GroupBox()
        Me.CheckBox43 = New System.Windows.Forms.CheckBox()
        Me.CheckBox44 = New System.Windows.Forms.CheckBox()
        Me.GroupBox27 = New System.Windows.Forms.GroupBox()
        Me.CheckBox45 = New System.Windows.Forms.CheckBox()
        Me.CheckBox46 = New System.Windows.Forms.CheckBox()
        Me.GroupBox28 = New System.Windows.Forms.GroupBox()
        Me.CheckBox47 = New System.Windows.Forms.CheckBox()
        Me.CheckBox48 = New System.Windows.Forms.CheckBox()
        Me.CheckBox49 = New System.Windows.Forms.CheckBox()
        Me.CheckBox50 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxEnergie = New System.Windows.Forms.GroupBox()
        Me.GroupBox39 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown9 = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox40 = New System.Windows.Forms.GroupBox()
        Me.CheckBox73 = New System.Windows.Forms.CheckBox()
        Me.CheckBox74 = New System.Windows.Forms.CheckBox()
        Me.CheckBox75 = New System.Windows.Forms.CheckBox()
        Me.GroupBox41 = New System.Windows.Forms.GroupBox()
        Me.CheckBox76 = New System.Windows.Forms.CheckBox()
        Me.CheckBox77 = New System.Windows.Forms.CheckBox()
        Me.GroupBox42 = New System.Windows.Forms.GroupBox()
        Me.CheckBox78 = New System.Windows.Forms.CheckBox()
        Me.CheckBox79 = New System.Windows.Forms.CheckBox()
        Me.GroupBox43 = New System.Windows.Forms.GroupBox()
        Me.CheckBox80 = New System.Windows.Forms.CheckBox()
        Me.CheckBox81 = New System.Windows.Forms.CheckBox()
        Me.CheckBox82 = New System.Windows.Forms.CheckBox()
        Me.CheckBox83 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxEquipementsMarchands = New System.Windows.Forms.GroupBox()
        Me.GroupBox29 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown7 = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox30 = New System.Windows.Forms.GroupBox()
        Me.CheckBox51 = New System.Windows.Forms.CheckBox()
        Me.CheckBox52 = New System.Windows.Forms.CheckBox()
        Me.CheckBox53 = New System.Windows.Forms.CheckBox()
        Me.GroupBox31 = New System.Windows.Forms.GroupBox()
        Me.CheckBox54 = New System.Windows.Forms.CheckBox()
        Me.CheckBox55 = New System.Windows.Forms.CheckBox()
        Me.GroupBox32 = New System.Windows.Forms.GroupBox()
        Me.CheckBox56 = New System.Windows.Forms.CheckBox()
        Me.CheckBox57 = New System.Windows.Forms.CheckBox()
        Me.GroupBox33 = New System.Windows.Forms.GroupBox()
        Me.CheckBox58 = New System.Windows.Forms.CheckBox()
        Me.CheckBox59 = New System.Windows.Forms.CheckBox()
        Me.CheckBox60 = New System.Windows.Forms.CheckBox()
        Me.CheckBox61 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxTransport = New System.Windows.Forms.GroupBox()
        Me.GroupBox34 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown8 = New System.Windows.Forms.NumericUpDown()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox35 = New System.Windows.Forms.GroupBox()
        Me.CheckBox62 = New System.Windows.Forms.CheckBox()
        Me.CheckBox63 = New System.Windows.Forms.CheckBox()
        Me.CheckBox64 = New System.Windows.Forms.CheckBox()
        Me.GroupBox36 = New System.Windows.Forms.GroupBox()
        Me.CheckBox65 = New System.Windows.Forms.CheckBox()
        Me.CheckBox66 = New System.Windows.Forms.CheckBox()
        Me.GroupBox37 = New System.Windows.Forms.GroupBox()
        Me.CheckBox67 = New System.Windows.Forms.CheckBox()
        Me.CheckBox68 = New System.Windows.Forms.CheckBox()
        Me.GroupBox38 = New System.Windows.Forms.GroupBox()
        Me.CheckBox69 = New System.Windows.Forms.CheckBox()
        Me.CheckBox70 = New System.Windows.Forms.CheckBox()
        Me.CheckBox71 = New System.Windows.Forms.CheckBox()
        Me.CheckBox72 = New System.Windows.Forms.CheckBox()
        Me.GroupBoxFiltresSpecifiquesToutInfras = New System.Windows.Forms.GroupBox()
        Me.ButtonSelectionnerToutInfra = New System.Windows.Forms.Button()
        Me.ButtonReinitialiserInfra = New System.Windows.Forms.Button()
        Me.CheckBoxAdministration = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEauAssainissement = New System.Windows.Forms.CheckBox()
        Me.CheckBoxElevage = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEquipementMarchand = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAgriculture = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEnergie = New System.Windows.Forms.CheckBox()
        Me.CheckBoxEducation = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSante = New System.Windows.Forms.CheckBox()
        Me.CheckBoxTransport = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSportsLoisirs = New System.Windows.Forms.CheckBox()
        Me.GroupBoxChoixInfras = New System.Windows.Forms.GroupBox()
        Me.LabelSourceEducation = New System.Windows.Forms.Label()
        Me.ButtonAfficherResultatSansFilter = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.OpenFileDialogEducation = New System.Windows.Forms.OpenFileDialog()
        Me.DataSetExcelko1 = New RIC_v3._03.DataSetExcelko1()
        Me.ButtonVoirRapportDetaille = New System.Windows.Forms.Button()
        Me.GroupBoxFiltresGeographiques.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBoxSelectFiltres.SuspendLayout()
        Me.GroupBoxEducation.SuspendLayout()
        Me.GroupBoxNombreDEleve.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxNiveauEnseigne.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBoxEtatInfraEducation.SuspendLayout()
        Me.GroupBoxAdministration.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBoxSante.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxDiplomeDuReferent.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox14.SuspendLayout()
        Me.GroupBoxEauEtAssainissement.SuspendLayout()
        Me.GroupBox15.SuspendLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxTypeDInstallation.SuspendLayout()
        Me.GroupBox17.SuspendLayout()
        Me.GroupBox18.SuspendLayout()
        Me.GroupBox19.SuspendLayout()
        Me.GroupBoxSportsEtLoisirs.SuspendLayout()
        Me.GroupBox16.SuspendLayout()
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox20.SuspendLayout()
        Me.GroupBox21.SuspendLayout()
        Me.GroupBox22.SuspendLayout()
        Me.GroupBox23.SuspendLayout()
        Me.GroupBoxAgricultureEtElevage.SuspendLayout()
        Me.GroupBox24.SuspendLayout()
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox25.SuspendLayout()
        Me.GroupBox26.SuspendLayout()
        Me.GroupBox27.SuspendLayout()
        Me.GroupBox28.SuspendLayout()
        Me.GroupBoxEnergie.SuspendLayout()
        Me.GroupBox39.SuspendLayout()
        CType(Me.NumericUpDown9, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox40.SuspendLayout()
        Me.GroupBox41.SuspendLayout()
        Me.GroupBox42.SuspendLayout()
        Me.GroupBox43.SuspendLayout()
        Me.GroupBoxEquipementsMarchands.SuspendLayout()
        Me.GroupBox29.SuspendLayout()
        CType(Me.NumericUpDown7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox30.SuspendLayout()
        Me.GroupBox31.SuspendLayout()
        Me.GroupBox32.SuspendLayout()
        Me.GroupBox33.SuspendLayout()
        Me.GroupBoxTransport.SuspendLayout()
        Me.GroupBox34.SuspendLayout()
        CType(Me.NumericUpDown8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox35.SuspendLayout()
        Me.GroupBox36.SuspendLayout()
        Me.GroupBox37.SuspendLayout()
        Me.GroupBox38.SuspendLayout()
        Me.GroupBoxFiltresSpecifiquesToutInfras.SuspendLayout()
        Me.GroupBoxChoixInfras.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSetExcelko1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TreeViewFiltresGeographiques
        '
        Me.TreeViewFiltresGeographiques.BackColor = System.Drawing.SystemColors.Control
        Me.TreeViewFiltresGeographiques.CheckBoxes = True
        Me.TreeViewFiltresGeographiques.Cursor = System.Windows.Forms.Cursors.Hand
        resources.ApplyResources(Me.TreeViewFiltresGeographiques, "TreeViewFiltresGeographiques")
        Me.TreeViewFiltresGeographiques.FullRowSelect = True
        Me.TreeViewFiltresGeographiques.HideSelection = False
        Me.TreeViewFiltresGeographiques.Name = "TreeViewFiltresGeographiques"
        Me.TreeViewFiltresGeographiques.Nodes.AddRange(New System.Windows.Forms.TreeNode() {CType(resources.GetObject("TreeViewFiltresGeographiques.Nodes"), System.Windows.Forms.TreeNode), CType(resources.GetObject("TreeViewFiltresGeographiques.Nodes1"), System.Windows.Forms.TreeNode), CType(resources.GetObject("TreeViewFiltresGeographiques.Nodes2"), System.Windows.Forms.TreeNode)})
        '
        'GroupBoxFiltresGeographiques
        '
        Me.GroupBoxFiltresGeographiques.Controls.Add(Me.TreeViewFiltresGeographiques)
        resources.ApplyResources(Me.GroupBoxFiltresGeographiques, "GroupBoxFiltresGeographiques")
        Me.GroupBoxFiltresGeographiques.Name = "GroupBoxFiltresGeographiques"
        Me.GroupBoxFiltresGeographiques.TabStop = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuToolStripMenuItem})
        resources.ApplyResources(Me.MenuStrip1, "MenuStrip1")
        Me.MenuStrip1.Name = "MenuStrip1"
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImporterSourceToolStripMenuItem, Me.AProposDeLoutilToolStripMenuItem, Me.QuitterToolStripMenuItem})
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        resources.ApplyResources(Me.MenuToolStripMenuItem, "MenuToolStripMenuItem")
        '
        'ImporterSourceToolStripMenuItem
        '
        Me.ImporterSourceToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EauEtAssainissementToolStripMenuItem, Me.SportsEtLoisirsToolStripMenuItem, Me.AdministrationToolStripMenuItem, Me.SantéToolStripMenuItem, Me.EducationToolStripMenuItem, Me.AgricultureToolStripMenuItem, Me.ElevageToolStripMenuItem, Me.TransportToolStripMenuItem, Me.EquipementsMarchandsToolStripMenuItem, Me.EnergieToolStripMenuItem})
        Me.ImporterSourceToolStripMenuItem.Name = "ImporterSourceToolStripMenuItem"
        resources.ApplyResources(Me.ImporterSourceToolStripMenuItem, "ImporterSourceToolStripMenuItem")
        '
        'EauEtAssainissementToolStripMenuItem
        '
        Me.EauEtAssainissementToolStripMenuItem.Name = "EauEtAssainissementToolStripMenuItem"
        resources.ApplyResources(Me.EauEtAssainissementToolStripMenuItem, "EauEtAssainissementToolStripMenuItem")
        '
        'SportsEtLoisirsToolStripMenuItem
        '
        Me.SportsEtLoisirsToolStripMenuItem.Name = "SportsEtLoisirsToolStripMenuItem"
        resources.ApplyResources(Me.SportsEtLoisirsToolStripMenuItem, "SportsEtLoisirsToolStripMenuItem")
        '
        'AdministrationToolStripMenuItem
        '
        Me.AdministrationToolStripMenuItem.Name = "AdministrationToolStripMenuItem"
        resources.ApplyResources(Me.AdministrationToolStripMenuItem, "AdministrationToolStripMenuItem")
        '
        'SantéToolStripMenuItem
        '
        Me.SantéToolStripMenuItem.Name = "SantéToolStripMenuItem"
        resources.ApplyResources(Me.SantéToolStripMenuItem, "SantéToolStripMenuItem")
        '
        'EducationToolStripMenuItem
        '
        Me.EducationToolStripMenuItem.Name = "EducationToolStripMenuItem"
        resources.ApplyResources(Me.EducationToolStripMenuItem, "EducationToolStripMenuItem")
        '
        'AgricultureToolStripMenuItem
        '
        Me.AgricultureToolStripMenuItem.Name = "AgricultureToolStripMenuItem"
        resources.ApplyResources(Me.AgricultureToolStripMenuItem, "AgricultureToolStripMenuItem")
        '
        'ElevageToolStripMenuItem
        '
        Me.ElevageToolStripMenuItem.Name = "ElevageToolStripMenuItem"
        resources.ApplyResources(Me.ElevageToolStripMenuItem, "ElevageToolStripMenuItem")
        '
        'TransportToolStripMenuItem
        '
        Me.TransportToolStripMenuItem.Name = "TransportToolStripMenuItem"
        resources.ApplyResources(Me.TransportToolStripMenuItem, "TransportToolStripMenuItem")
        '
        'EquipementsMarchandsToolStripMenuItem
        '
        Me.EquipementsMarchandsToolStripMenuItem.Name = "EquipementsMarchandsToolStripMenuItem"
        resources.ApplyResources(Me.EquipementsMarchandsToolStripMenuItem, "EquipementsMarchandsToolStripMenuItem")
        '
        'EnergieToolStripMenuItem
        '
        Me.EnergieToolStripMenuItem.Name = "EnergieToolStripMenuItem"
        resources.ApplyResources(Me.EnergieToolStripMenuItem, "EnergieToolStripMenuItem")
        '
        'AProposDeLoutilToolStripMenuItem
        '
        Me.AProposDeLoutilToolStripMenuItem.Name = "AProposDeLoutilToolStripMenuItem"
        resources.ApplyResources(Me.AProposDeLoutilToolStripMenuItem, "AProposDeLoutilToolStripMenuItem")
        '
        'QuitterToolStripMenuItem
        '
        Me.QuitterToolStripMenuItem.Name = "QuitterToolStripMenuItem"
        resources.ApplyResources(Me.QuitterToolStripMenuItem, "QuitterToolStripMenuItem")
        '
        'GroupBoxSelectFiltres
        '
        Me.GroupBoxSelectFiltres.Controls.Add(Me.ButtonCancelFiltres)
        Me.GroupBoxSelectFiltres.Controls.Add(Me.ButtonSelectAllFiltres)
        Me.GroupBoxSelectFiltres.Controls.Add(Me.ButtonAppliquerFiltres)
        Me.GroupBoxSelectFiltres.Controls.Add(Me.ComboBoxInfraAFiltrer)
        resources.ApplyResources(Me.GroupBoxSelectFiltres, "GroupBoxSelectFiltres")
        Me.GroupBoxSelectFiltres.Name = "GroupBoxSelectFiltres"
        Me.GroupBoxSelectFiltres.TabStop = False
        '
        'ButtonCancelFiltres
        '
        resources.ApplyResources(Me.ButtonCancelFiltres, "ButtonCancelFiltres")
        Me.ButtonCancelFiltres.Name = "ButtonCancelFiltres"
        Me.ButtonCancelFiltres.UseVisualStyleBackColor = True
        '
        'ButtonSelectAllFiltres
        '
        resources.ApplyResources(Me.ButtonSelectAllFiltres, "ButtonSelectAllFiltres")
        Me.ButtonSelectAllFiltres.Name = "ButtonSelectAllFiltres"
        Me.ButtonSelectAllFiltres.UseVisualStyleBackColor = True
        '
        'ButtonAppliquerFiltres
        '
        resources.ApplyResources(Me.ButtonAppliquerFiltres, "ButtonAppliquerFiltres")
        Me.ButtonAppliquerFiltres.Name = "ButtonAppliquerFiltres"
        Me.ButtonAppliquerFiltres.UseVisualStyleBackColor = True
        '
        'ComboBoxInfraAFiltrer
        '
        resources.ApplyResources(Me.ComboBoxInfraAFiltrer, "ComboBoxInfraAFiltrer")
        Me.ComboBoxInfraAFiltrer.FormattingEnabled = True
        Me.ComboBoxInfraAFiltrer.Name = "ComboBoxInfraAFiltrer"
        '
        'GroupBoxEducation
        '
        resources.ApplyResources(Me.GroupBoxEducation, "GroupBoxEducation")
        Me.GroupBoxEducation.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.GroupBoxEducation.Controls.Add(Me.GroupBoxNombreDEleve)
        Me.GroupBoxEducation.Controls.Add(Me.GroupBoxNiveauEnseigne)
        Me.GroupBoxEducation.Controls.Add(Me.GroupBox3)
        Me.GroupBoxEducation.Controls.Add(Me.GroupBoxEtatInfraEducation)
        Me.GroupBoxEducation.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxEducation.Name = "GroupBoxEducation"
        Me.GroupBoxEducation.TabStop = False
        '
        'GroupBoxNombreDEleve
        '
        Me.GroupBoxNombreDEleve.Controls.Add(Me.NumericUpDown1)
        Me.GroupBoxNombreDEleve.Controls.Add(Me.LabelNombreEleveSuperieurA)
        resources.ApplyResources(Me.GroupBoxNombreDEleve, "GroupBoxNombreDEleve")
        Me.GroupBoxNombreDEleve.Name = "GroupBoxNombreDEleve"
        Me.GroupBoxNombreDEleve.TabStop = False
        '
        'NumericUpDown1
        '
        resources.ApplyResources(Me.NumericUpDown1, "NumericUpDown1")
        Me.NumericUpDown1.Name = "NumericUpDown1"
        '
        'LabelNombreEleveSuperieurA
        '
        resources.ApplyResources(Me.LabelNombreEleveSuperieurA, "LabelNombreEleveSuperieurA")
        Me.LabelNombreEleveSuperieurA.Name = "LabelNombreEleveSuperieurA"
        '
        'GroupBoxNiveauEnseigne
        '
        Me.GroupBoxNiveauEnseigne.Controls.Add(Me.CheckBoxLycee)
        Me.GroupBoxNiveauEnseigne.Controls.Add(Me.CheckBoxCollege)
        Me.GroupBoxNiveauEnseigne.Controls.Add(Me.CheckBoxPrimaire)
        Me.GroupBoxNiveauEnseigne.Controls.Add(Me.CheckBoxPrescolaire)
        resources.ApplyResources(Me.GroupBoxNiveauEnseigne, "GroupBoxNiveauEnseigne")
        Me.GroupBoxNiveauEnseigne.Name = "GroupBoxNiveauEnseigne"
        Me.GroupBoxNiveauEnseigne.TabStop = False
        '
        'CheckBoxLycee
        '
        resources.ApplyResources(Me.CheckBoxLycee, "CheckBoxLycee")
        Me.CheckBoxLycee.Name = "CheckBoxLycee"
        Me.CheckBoxLycee.UseVisualStyleBackColor = True
        '
        'CheckBoxCollege
        '
        resources.ApplyResources(Me.CheckBoxCollege, "CheckBoxCollege")
        Me.CheckBoxCollege.Name = "CheckBoxCollege"
        Me.CheckBoxCollege.UseVisualStyleBackColor = True
        '
        'CheckBoxPrimaire
        '
        resources.ApplyResources(Me.CheckBoxPrimaire, "CheckBoxPrimaire")
        Me.CheckBoxPrimaire.Name = "CheckBoxPrimaire"
        Me.CheckBoxPrimaire.UseVisualStyleBackColor = True
        '
        'CheckBoxPrescolaire
        '
        resources.ApplyResources(Me.CheckBoxPrescolaire, "CheckBoxPrescolaire")
        Me.CheckBoxPrescolaire.Name = "CheckBoxPrescolaire"
        Me.CheckBoxPrescolaire.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CheckBoxEcolePublique)
        Me.GroupBox3.Controls.Add(Me.CheckBoxEcolePrive)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'CheckBoxEcolePublique
        '
        resources.ApplyResources(Me.CheckBoxEcolePublique, "CheckBoxEcolePublique")
        Me.CheckBoxEcolePublique.Name = "CheckBoxEcolePublique"
        Me.CheckBoxEcolePublique.UseVisualStyleBackColor = True
        '
        'CheckBoxEcolePrive
        '
        resources.ApplyResources(Me.CheckBoxEcolePrive, "CheckBoxEcolePrive")
        Me.CheckBoxEcolePrive.Name = "CheckBoxEcolePrive"
        Me.CheckBoxEcolePrive.UseVisualStyleBackColor = True
        '
        'GroupBoxEtatInfraEducation
        '
        Me.GroupBoxEtatInfraEducation.Controls.Add(Me.CheckBoxInutilisableEducation)
        Me.GroupBoxEtatInfraEducation.Controls.Add(Me.CheckBox3)
        Me.GroupBoxEtatInfraEducation.Controls.Add(Me.CheckBoxEtatUsageEducation)
        Me.GroupBoxEtatInfraEducation.Controls.Add(Me.CheckBoxBonEtatEducation)
        resources.ApplyResources(Me.GroupBoxEtatInfraEducation, "GroupBoxEtatInfraEducation")
        Me.GroupBoxEtatInfraEducation.Name = "GroupBoxEtatInfraEducation"
        Me.GroupBoxEtatInfraEducation.TabStop = False
        '
        'CheckBoxInutilisableEducation
        '
        resources.ApplyResources(Me.CheckBoxInutilisableEducation, "CheckBoxInutilisableEducation")
        Me.CheckBoxInutilisableEducation.Name = "CheckBoxInutilisableEducation"
        Me.CheckBoxInutilisableEducation.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        resources.ApplyResources(Me.CheckBox3, "CheckBox3")
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBoxEtatUsageEducation
        '
        resources.ApplyResources(Me.CheckBoxEtatUsageEducation, "CheckBoxEtatUsageEducation")
        Me.CheckBoxEtatUsageEducation.Name = "CheckBoxEtatUsageEducation"
        Me.CheckBoxEtatUsageEducation.UseVisualStyleBackColor = True
        '
        'CheckBoxBonEtatEducation
        '
        resources.ApplyResources(Me.CheckBoxBonEtatEducation, "CheckBoxBonEtatEducation")
        Me.CheckBoxBonEtatEducation.Name = "CheckBoxBonEtatEducation"
        Me.CheckBoxBonEtatEducation.UseVisualStyleBackColor = True
        '
        'GroupBoxAdministration
        '
        resources.ApplyResources(Me.GroupBoxAdministration, "GroupBoxAdministration")
        Me.GroupBoxAdministration.Controls.Add(Me.GroupBox6)
        Me.GroupBoxAdministration.Controls.Add(Me.GroupBox8)
        Me.GroupBoxAdministration.Controls.Add(Me.GroupBox9)
        Me.GroupBoxAdministration.Controls.Add(Me.GroupBox10)
        Me.GroupBoxAdministration.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxAdministration.Name = "GroupBoxAdministration"
        Me.GroupBoxAdministration.TabStop = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox6.Controls.Add(Me.Label1)
        resources.ApplyResources(Me.GroupBox6, "GroupBox6")
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.TabStop = False
        '
        'NumericUpDown2
        '
        resources.ApplyResources(Me.NumericUpDown2, "NumericUpDown2")
        Me.NumericUpDown2.Name = "NumericUpDown2"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.CheckBox1)
        Me.GroupBox8.Controls.Add(Me.CheckBox2)
        Me.GroupBox8.Controls.Add(Me.CheckBox4)
        Me.GroupBox8.Controls.Add(Me.CheckBox5)
        resources.ApplyResources(Me.GroupBox8, "GroupBox8")
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.TabStop = False
        '
        'CheckBox1
        '
        resources.ApplyResources(Me.CheckBox1, "CheckBox1")
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        resources.ApplyResources(Me.CheckBox2, "CheckBox2")
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        resources.ApplyResources(Me.CheckBox4, "CheckBox4")
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        resources.ApplyResources(Me.CheckBox5, "CheckBox5")
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.CheckBox6)
        Me.GroupBox9.Controls.Add(Me.CheckBox7)
        resources.ApplyResources(Me.GroupBox9, "GroupBox9")
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.TabStop = False
        '
        'CheckBox6
        '
        resources.ApplyResources(Me.CheckBox6, "CheckBox6")
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        resources.ApplyResources(Me.CheckBox7, "CheckBox7")
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.CheckBox8)
        Me.GroupBox10.Controls.Add(Me.CheckBox9)
        Me.GroupBox10.Controls.Add(Me.CheckBox10)
        Me.GroupBox10.Controls.Add(Me.CheckBox14)
        resources.ApplyResources(Me.GroupBox10, "GroupBox10")
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.TabStop = False
        '
        'CheckBox8
        '
        resources.ApplyResources(Me.CheckBox8, "CheckBox8")
        Me.CheckBox8.Name = "CheckBox8"
        Me.CheckBox8.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        resources.ApplyResources(Me.CheckBox9, "CheckBox9")
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'CheckBox10
        '
        resources.ApplyResources(Me.CheckBox10, "CheckBox10")
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.UseVisualStyleBackColor = True
        '
        'CheckBox14
        '
        resources.ApplyResources(Me.CheckBox14, "CheckBox14")
        Me.CheckBox14.Name = "CheckBox14"
        Me.CheckBox14.UseVisualStyleBackColor = True
        '
        'GroupBoxSante
        '
        resources.ApplyResources(Me.GroupBoxSante, "GroupBoxSante")
        Me.GroupBoxSante.Controls.Add(Me.GroupBox11)
        Me.GroupBoxSante.Controls.Add(Me.GroupBoxDiplomeDuReferent)
        Me.GroupBoxSante.Controls.Add(Me.GroupBox12)
        Me.GroupBoxSante.Controls.Add(Me.GroupBox13)
        Me.GroupBoxSante.Controls.Add(Me.GroupBox14)
        Me.GroupBoxSante.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxSante.Name = "GroupBoxSante"
        Me.GroupBoxSante.TabStop = False
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.NumericUpDown3)
        Me.GroupBox11.Controls.Add(Me.Label2)
        resources.ApplyResources(Me.GroupBox11, "GroupBox11")
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.TabStop = False
        '
        'NumericUpDown3
        '
        resources.ApplyResources(Me.NumericUpDown3, "NumericUpDown3")
        Me.NumericUpDown3.Name = "NumericUpDown3"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'GroupBoxDiplomeDuReferent
        '
        Me.GroupBoxDiplomeDuReferent.Controls.Add(Me.CheckBoxInfirmier)
        Me.GroupBoxDiplomeDuReferent.Controls.Add(Me.CheckBoxSageFemme)
        Me.GroupBoxDiplomeDuReferent.Controls.Add(Me.CheckBoxDocteur)
        resources.ApplyResources(Me.GroupBoxDiplomeDuReferent, "GroupBoxDiplomeDuReferent")
        Me.GroupBoxDiplomeDuReferent.Name = "GroupBoxDiplomeDuReferent"
        Me.GroupBoxDiplomeDuReferent.TabStop = False
        '
        'CheckBoxInfirmier
        '
        resources.ApplyResources(Me.CheckBoxInfirmier, "CheckBoxInfirmier")
        Me.CheckBoxInfirmier.Name = "CheckBoxInfirmier"
        Me.CheckBoxInfirmier.UseVisualStyleBackColor = True
        '
        'CheckBoxSageFemme
        '
        resources.ApplyResources(Me.CheckBoxSageFemme, "CheckBoxSageFemme")
        Me.CheckBoxSageFemme.Name = "CheckBoxSageFemme"
        Me.CheckBoxSageFemme.UseVisualStyleBackColor = True
        '
        'CheckBoxDocteur
        '
        resources.ApplyResources(Me.CheckBoxDocteur, "CheckBoxDocteur")
        Me.CheckBoxDocteur.Name = "CheckBoxDocteur"
        Me.CheckBoxDocteur.UseVisualStyleBackColor = True
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.CheckBoxCSBII)
        Me.GroupBox12.Controls.Add(Me.CheckBoxCSBI)
        resources.ApplyResources(Me.GroupBox12, "GroupBox12")
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.TabStop = False
        '
        'CheckBoxCSBII
        '
        resources.ApplyResources(Me.CheckBoxCSBII, "CheckBoxCSBII")
        Me.CheckBoxCSBII.Name = "CheckBoxCSBII"
        Me.CheckBoxCSBII.UseVisualStyleBackColor = True
        '
        'CheckBoxCSBI
        '
        resources.ApplyResources(Me.CheckBoxCSBI, "CheckBoxCSBI")
        Me.CheckBoxCSBI.Name = "CheckBoxCSBI"
        Me.CheckBoxCSBI.UseVisualStyleBackColor = True
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.CheckBox19)
        Me.GroupBox13.Controls.Add(Me.CheckBox20)
        resources.ApplyResources(Me.GroupBox13, "GroupBox13")
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.TabStop = False
        '
        'CheckBox19
        '
        resources.ApplyResources(Me.CheckBox19, "CheckBox19")
        Me.CheckBox19.Name = "CheckBox19"
        Me.CheckBox19.UseVisualStyleBackColor = True
        '
        'CheckBox20
        '
        resources.ApplyResources(Me.CheckBox20, "CheckBox20")
        Me.CheckBox20.Name = "CheckBox20"
        Me.CheckBox20.UseVisualStyleBackColor = True
        '
        'GroupBox14
        '
        Me.GroupBox14.Controls.Add(Me.CheckBox21)
        Me.GroupBox14.Controls.Add(Me.CheckBox22)
        Me.GroupBox14.Controls.Add(Me.CheckBox23)
        Me.GroupBox14.Controls.Add(Me.CheckBox24)
        resources.ApplyResources(Me.GroupBox14, "GroupBox14")
        Me.GroupBox14.Name = "GroupBox14"
        Me.GroupBox14.TabStop = False
        '
        'CheckBox21
        '
        resources.ApplyResources(Me.CheckBox21, "CheckBox21")
        Me.CheckBox21.Name = "CheckBox21"
        Me.CheckBox21.UseVisualStyleBackColor = True
        '
        'CheckBox22
        '
        resources.ApplyResources(Me.CheckBox22, "CheckBox22")
        Me.CheckBox22.Name = "CheckBox22"
        Me.CheckBox22.UseVisualStyleBackColor = True
        '
        'CheckBox23
        '
        resources.ApplyResources(Me.CheckBox23, "CheckBox23")
        Me.CheckBox23.Name = "CheckBox23"
        Me.CheckBox23.UseVisualStyleBackColor = True
        '
        'CheckBox24
        '
        resources.ApplyResources(Me.CheckBox24, "CheckBox24")
        Me.CheckBox24.Name = "CheckBox24"
        Me.CheckBox24.UseVisualStyleBackColor = True
        '
        'GroupBoxEauEtAssainissement
        '
        resources.ApplyResources(Me.GroupBoxEauEtAssainissement, "GroupBoxEauEtAssainissement")
        Me.GroupBoxEauEtAssainissement.BackColor = System.Drawing.Color.LavenderBlush
        Me.GroupBoxEauEtAssainissement.Controls.Add(Me.GroupBox15)
        Me.GroupBoxEauEtAssainissement.Controls.Add(Me.GroupBoxTypeDInstallation)
        Me.GroupBoxEauEtAssainissement.Controls.Add(Me.GroupBox17)
        Me.GroupBoxEauEtAssainissement.Controls.Add(Me.GroupBox18)
        Me.GroupBoxEauEtAssainissement.Controls.Add(Me.GroupBox19)
        Me.GroupBoxEauEtAssainissement.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxEauEtAssainissement.Name = "GroupBoxEauEtAssainissement"
        Me.GroupBoxEauEtAssainissement.TabStop = False
        '
        'GroupBox15
        '
        Me.GroupBox15.Controls.Add(Me.NumericUpDown4)
        Me.GroupBox15.Controls.Add(Me.Label3)
        resources.ApplyResources(Me.GroupBox15, "GroupBox15")
        Me.GroupBox15.Name = "GroupBox15"
        Me.GroupBox15.TabStop = False
        '
        'NumericUpDown4
        '
        resources.ApplyResources(Me.NumericUpDown4, "NumericUpDown4")
        Me.NumericUpDown4.Name = "NumericUpDown4"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'GroupBoxTypeDInstallation
        '
        Me.GroupBoxTypeDInstallation.Controls.Add(Me.CheckBoxChâteauDEau)
        Me.GroupBoxTypeDInstallation.Controls.Add(Me.CheckBoxPuits)
        Me.GroupBoxTypeDInstallation.Controls.Add(Me.CheckBoxBorneFontaine)
        resources.ApplyResources(Me.GroupBoxTypeDInstallation, "GroupBoxTypeDInstallation")
        Me.GroupBoxTypeDInstallation.Name = "GroupBoxTypeDInstallation"
        Me.GroupBoxTypeDInstallation.TabStop = False
        '
        'CheckBoxChâteauDEau
        '
        resources.ApplyResources(Me.CheckBoxChâteauDEau, "CheckBoxChâteauDEau")
        Me.CheckBoxChâteauDEau.Name = "CheckBoxChâteauDEau"
        Me.CheckBoxChâteauDEau.UseVisualStyleBackColor = True
        '
        'CheckBoxPuits
        '
        resources.ApplyResources(Me.CheckBoxPuits, "CheckBoxPuits")
        Me.CheckBoxPuits.Name = "CheckBoxPuits"
        Me.CheckBoxPuits.UseVisualStyleBackColor = True
        '
        'CheckBoxBorneFontaine
        '
        resources.ApplyResources(Me.CheckBoxBorneFontaine, "CheckBoxBorneFontaine")
        Me.CheckBoxBorneFontaine.Name = "CheckBoxBorneFontaine"
        Me.CheckBoxBorneFontaine.UseVisualStyleBackColor = True
        '
        'GroupBox17
        '
        Me.GroupBox17.Controls.Add(Me.CheckBox18)
        Me.GroupBox17.Controls.Add(Me.CheckBox25)
        resources.ApplyResources(Me.GroupBox17, "GroupBox17")
        Me.GroupBox17.Name = "GroupBox17"
        Me.GroupBox17.TabStop = False
        '
        'CheckBox18
        '
        resources.ApplyResources(Me.CheckBox18, "CheckBox18")
        Me.CheckBox18.Name = "CheckBox18"
        Me.CheckBox18.UseVisualStyleBackColor = True
        '
        'CheckBox25
        '
        resources.ApplyResources(Me.CheckBox25, "CheckBox25")
        Me.CheckBox25.Name = "CheckBox25"
        Me.CheckBox25.UseVisualStyleBackColor = True
        '
        'GroupBox18
        '
        Me.GroupBox18.Controls.Add(Me.CheckBox26)
        Me.GroupBox18.Controls.Add(Me.CheckBox27)
        resources.ApplyResources(Me.GroupBox18, "GroupBox18")
        Me.GroupBox18.Name = "GroupBox18"
        Me.GroupBox18.TabStop = False
        '
        'CheckBox26
        '
        resources.ApplyResources(Me.CheckBox26, "CheckBox26")
        Me.CheckBox26.Name = "CheckBox26"
        Me.CheckBox26.UseVisualStyleBackColor = True
        '
        'CheckBox27
        '
        resources.ApplyResources(Me.CheckBox27, "CheckBox27")
        Me.CheckBox27.Name = "CheckBox27"
        Me.CheckBox27.UseVisualStyleBackColor = True
        '
        'GroupBox19
        '
        Me.GroupBox19.Controls.Add(Me.CheckBox28)
        Me.GroupBox19.Controls.Add(Me.CheckBox29)
        Me.GroupBox19.Controls.Add(Me.CheckBox30)
        Me.GroupBox19.Controls.Add(Me.CheckBox31)
        resources.ApplyResources(Me.GroupBox19, "GroupBox19")
        Me.GroupBox19.Name = "GroupBox19"
        Me.GroupBox19.TabStop = False
        '
        'CheckBox28
        '
        resources.ApplyResources(Me.CheckBox28, "CheckBox28")
        Me.CheckBox28.Name = "CheckBox28"
        Me.CheckBox28.UseVisualStyleBackColor = True
        '
        'CheckBox29
        '
        resources.ApplyResources(Me.CheckBox29, "CheckBox29")
        Me.CheckBox29.Name = "CheckBox29"
        Me.CheckBox29.UseVisualStyleBackColor = True
        '
        'CheckBox30
        '
        resources.ApplyResources(Me.CheckBox30, "CheckBox30")
        Me.CheckBox30.Name = "CheckBox30"
        Me.CheckBox30.UseVisualStyleBackColor = True
        '
        'CheckBox31
        '
        resources.ApplyResources(Me.CheckBox31, "CheckBox31")
        Me.CheckBox31.Name = "CheckBox31"
        Me.CheckBox31.UseVisualStyleBackColor = True
        '
        'GroupBoxSportsEtLoisirs
        '
        resources.ApplyResources(Me.GroupBoxSportsEtLoisirs, "GroupBoxSportsEtLoisirs")
        Me.GroupBoxSportsEtLoisirs.BackColor = System.Drawing.Color.Magenta
        Me.GroupBoxSportsEtLoisirs.Controls.Add(Me.GroupBox16)
        Me.GroupBoxSportsEtLoisirs.Controls.Add(Me.GroupBox20)
        Me.GroupBoxSportsEtLoisirs.Controls.Add(Me.GroupBox21)
        Me.GroupBoxSportsEtLoisirs.Controls.Add(Me.GroupBox22)
        Me.GroupBoxSportsEtLoisirs.Controls.Add(Me.GroupBox23)
        Me.GroupBoxSportsEtLoisirs.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxSportsEtLoisirs.Name = "GroupBoxSportsEtLoisirs"
        Me.GroupBoxSportsEtLoisirs.TabStop = False
        '
        'GroupBox16
        '
        Me.GroupBox16.Controls.Add(Me.NumericUpDown5)
        Me.GroupBox16.Controls.Add(Me.Label4)
        resources.ApplyResources(Me.GroupBox16, "GroupBox16")
        Me.GroupBox16.Name = "GroupBox16"
        Me.GroupBox16.TabStop = False
        '
        'NumericUpDown5
        '
        resources.ApplyResources(Me.NumericUpDown5, "NumericUpDown5")
        Me.NumericUpDown5.Name = "NumericUpDown5"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'GroupBox20
        '
        Me.GroupBox20.Controls.Add(Me.CheckBox15)
        Me.GroupBox20.Controls.Add(Me.CheckBox16)
        Me.GroupBox20.Controls.Add(Me.CheckBox17)
        resources.ApplyResources(Me.GroupBox20, "GroupBox20")
        Me.GroupBox20.Name = "GroupBox20"
        Me.GroupBox20.TabStop = False
        '
        'CheckBox15
        '
        resources.ApplyResources(Me.CheckBox15, "CheckBox15")
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.UseVisualStyleBackColor = True
        '
        'CheckBox16
        '
        resources.ApplyResources(Me.CheckBox16, "CheckBox16")
        Me.CheckBox16.Name = "CheckBox16"
        Me.CheckBox16.UseVisualStyleBackColor = True
        '
        'CheckBox17
        '
        resources.ApplyResources(Me.CheckBox17, "CheckBox17")
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.UseVisualStyleBackColor = True
        '
        'GroupBox21
        '
        Me.GroupBox21.Controls.Add(Me.CheckBox32)
        Me.GroupBox21.Controls.Add(Me.CheckBox33)
        resources.ApplyResources(Me.GroupBox21, "GroupBox21")
        Me.GroupBox21.Name = "GroupBox21"
        Me.GroupBox21.TabStop = False
        '
        'CheckBox32
        '
        resources.ApplyResources(Me.CheckBox32, "CheckBox32")
        Me.CheckBox32.Name = "CheckBox32"
        Me.CheckBox32.UseVisualStyleBackColor = True
        '
        'CheckBox33
        '
        resources.ApplyResources(Me.CheckBox33, "CheckBox33")
        Me.CheckBox33.Name = "CheckBox33"
        Me.CheckBox33.UseVisualStyleBackColor = True
        '
        'GroupBox22
        '
        Me.GroupBox22.Controls.Add(Me.CheckBox34)
        Me.GroupBox22.Controls.Add(Me.CheckBox35)
        resources.ApplyResources(Me.GroupBox22, "GroupBox22")
        Me.GroupBox22.Name = "GroupBox22"
        Me.GroupBox22.TabStop = False
        '
        'CheckBox34
        '
        resources.ApplyResources(Me.CheckBox34, "CheckBox34")
        Me.CheckBox34.Name = "CheckBox34"
        Me.CheckBox34.UseVisualStyleBackColor = True
        '
        'CheckBox35
        '
        resources.ApplyResources(Me.CheckBox35, "CheckBox35")
        Me.CheckBox35.Name = "CheckBox35"
        Me.CheckBox35.UseVisualStyleBackColor = True
        '
        'GroupBox23
        '
        Me.GroupBox23.Controls.Add(Me.CheckBox36)
        Me.GroupBox23.Controls.Add(Me.CheckBox37)
        Me.GroupBox23.Controls.Add(Me.CheckBox38)
        Me.GroupBox23.Controls.Add(Me.CheckBox39)
        resources.ApplyResources(Me.GroupBox23, "GroupBox23")
        Me.GroupBox23.Name = "GroupBox23"
        Me.GroupBox23.TabStop = False
        '
        'CheckBox36
        '
        resources.ApplyResources(Me.CheckBox36, "CheckBox36")
        Me.CheckBox36.Name = "CheckBox36"
        Me.CheckBox36.UseVisualStyleBackColor = True
        '
        'CheckBox37
        '
        resources.ApplyResources(Me.CheckBox37, "CheckBox37")
        Me.CheckBox37.Name = "CheckBox37"
        Me.CheckBox37.UseVisualStyleBackColor = True
        '
        'CheckBox38
        '
        resources.ApplyResources(Me.CheckBox38, "CheckBox38")
        Me.CheckBox38.Name = "CheckBox38"
        Me.CheckBox38.UseVisualStyleBackColor = True
        '
        'CheckBox39
        '
        resources.ApplyResources(Me.CheckBox39, "CheckBox39")
        Me.CheckBox39.Name = "CheckBox39"
        Me.CheckBox39.UseVisualStyleBackColor = True
        '
        'GroupBoxAgricultureEtElevage
        '
        resources.ApplyResources(Me.GroupBoxAgricultureEtElevage, "GroupBoxAgricultureEtElevage")
        Me.GroupBoxAgricultureEtElevage.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBoxAgricultureEtElevage.Controls.Add(Me.GroupBox24)
        Me.GroupBoxAgricultureEtElevage.Controls.Add(Me.GroupBox25)
        Me.GroupBoxAgricultureEtElevage.Controls.Add(Me.GroupBox26)
        Me.GroupBoxAgricultureEtElevage.Controls.Add(Me.GroupBox27)
        Me.GroupBoxAgricultureEtElevage.Controls.Add(Me.GroupBox28)
        Me.GroupBoxAgricultureEtElevage.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxAgricultureEtElevage.Name = "GroupBoxAgricultureEtElevage"
        Me.GroupBoxAgricultureEtElevage.TabStop = False
        '
        'GroupBox24
        '
        Me.GroupBox24.Controls.Add(Me.NumericUpDown6)
        Me.GroupBox24.Controls.Add(Me.Label5)
        resources.ApplyResources(Me.GroupBox24, "GroupBox24")
        Me.GroupBox24.Name = "GroupBox24"
        Me.GroupBox24.TabStop = False
        '
        'NumericUpDown6
        '
        resources.ApplyResources(Me.NumericUpDown6, "NumericUpDown6")
        Me.NumericUpDown6.Name = "NumericUpDown6"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'GroupBox25
        '
        Me.GroupBox25.Controls.Add(Me.CheckBox40)
        Me.GroupBox25.Controls.Add(Me.CheckBox41)
        Me.GroupBox25.Controls.Add(Me.CheckBox42)
        resources.ApplyResources(Me.GroupBox25, "GroupBox25")
        Me.GroupBox25.Name = "GroupBox25"
        Me.GroupBox25.TabStop = False
        '
        'CheckBox40
        '
        resources.ApplyResources(Me.CheckBox40, "CheckBox40")
        Me.CheckBox40.Name = "CheckBox40"
        Me.CheckBox40.UseVisualStyleBackColor = True
        '
        'CheckBox41
        '
        resources.ApplyResources(Me.CheckBox41, "CheckBox41")
        Me.CheckBox41.Name = "CheckBox41"
        Me.CheckBox41.UseVisualStyleBackColor = True
        '
        'CheckBox42
        '
        resources.ApplyResources(Me.CheckBox42, "CheckBox42")
        Me.CheckBox42.Name = "CheckBox42"
        Me.CheckBox42.UseVisualStyleBackColor = True
        '
        'GroupBox26
        '
        Me.GroupBox26.Controls.Add(Me.CheckBox43)
        Me.GroupBox26.Controls.Add(Me.CheckBox44)
        resources.ApplyResources(Me.GroupBox26, "GroupBox26")
        Me.GroupBox26.Name = "GroupBox26"
        Me.GroupBox26.TabStop = False
        '
        'CheckBox43
        '
        resources.ApplyResources(Me.CheckBox43, "CheckBox43")
        Me.CheckBox43.Name = "CheckBox43"
        Me.CheckBox43.UseVisualStyleBackColor = True
        '
        'CheckBox44
        '
        resources.ApplyResources(Me.CheckBox44, "CheckBox44")
        Me.CheckBox44.Name = "CheckBox44"
        Me.CheckBox44.UseVisualStyleBackColor = True
        '
        'GroupBox27
        '
        Me.GroupBox27.Controls.Add(Me.CheckBox45)
        Me.GroupBox27.Controls.Add(Me.CheckBox46)
        resources.ApplyResources(Me.GroupBox27, "GroupBox27")
        Me.GroupBox27.Name = "GroupBox27"
        Me.GroupBox27.TabStop = False
        '
        'CheckBox45
        '
        resources.ApplyResources(Me.CheckBox45, "CheckBox45")
        Me.CheckBox45.Name = "CheckBox45"
        Me.CheckBox45.UseVisualStyleBackColor = True
        '
        'CheckBox46
        '
        resources.ApplyResources(Me.CheckBox46, "CheckBox46")
        Me.CheckBox46.Name = "CheckBox46"
        Me.CheckBox46.UseVisualStyleBackColor = True
        '
        'GroupBox28
        '
        Me.GroupBox28.Controls.Add(Me.CheckBox47)
        Me.GroupBox28.Controls.Add(Me.CheckBox48)
        Me.GroupBox28.Controls.Add(Me.CheckBox49)
        Me.GroupBox28.Controls.Add(Me.CheckBox50)
        resources.ApplyResources(Me.GroupBox28, "GroupBox28")
        Me.GroupBox28.Name = "GroupBox28"
        Me.GroupBox28.TabStop = False
        '
        'CheckBox47
        '
        resources.ApplyResources(Me.CheckBox47, "CheckBox47")
        Me.CheckBox47.Name = "CheckBox47"
        Me.CheckBox47.UseVisualStyleBackColor = True
        '
        'CheckBox48
        '
        resources.ApplyResources(Me.CheckBox48, "CheckBox48")
        Me.CheckBox48.Name = "CheckBox48"
        Me.CheckBox48.UseVisualStyleBackColor = True
        '
        'CheckBox49
        '
        resources.ApplyResources(Me.CheckBox49, "CheckBox49")
        Me.CheckBox49.Name = "CheckBox49"
        Me.CheckBox49.UseVisualStyleBackColor = True
        '
        'CheckBox50
        '
        resources.ApplyResources(Me.CheckBox50, "CheckBox50")
        Me.CheckBox50.Name = "CheckBox50"
        Me.CheckBox50.UseVisualStyleBackColor = True
        '
        'GroupBoxEnergie
        '
        resources.ApplyResources(Me.GroupBoxEnergie, "GroupBoxEnergie")
        Me.GroupBoxEnergie.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBoxEnergie.Controls.Add(Me.GroupBox39)
        Me.GroupBoxEnergie.Controls.Add(Me.GroupBox40)
        Me.GroupBoxEnergie.Controls.Add(Me.GroupBox41)
        Me.GroupBoxEnergie.Controls.Add(Me.GroupBox42)
        Me.GroupBoxEnergie.Controls.Add(Me.GroupBox43)
        Me.GroupBoxEnergie.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxEnergie.Name = "GroupBoxEnergie"
        Me.GroupBoxEnergie.TabStop = False
        '
        'GroupBox39
        '
        Me.GroupBox39.Controls.Add(Me.NumericUpDown9)
        Me.GroupBox39.Controls.Add(Me.Label8)
        resources.ApplyResources(Me.GroupBox39, "GroupBox39")
        Me.GroupBox39.Name = "GroupBox39"
        Me.GroupBox39.TabStop = False
        '
        'NumericUpDown9
        '
        resources.ApplyResources(Me.NumericUpDown9, "NumericUpDown9")
        Me.NumericUpDown9.Name = "NumericUpDown9"
        '
        'Label8
        '
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'GroupBox40
        '
        Me.GroupBox40.Controls.Add(Me.CheckBox73)
        Me.GroupBox40.Controls.Add(Me.CheckBox74)
        Me.GroupBox40.Controls.Add(Me.CheckBox75)
        resources.ApplyResources(Me.GroupBox40, "GroupBox40")
        Me.GroupBox40.Name = "GroupBox40"
        Me.GroupBox40.TabStop = False
        '
        'CheckBox73
        '
        resources.ApplyResources(Me.CheckBox73, "CheckBox73")
        Me.CheckBox73.Name = "CheckBox73"
        Me.CheckBox73.UseVisualStyleBackColor = True
        '
        'CheckBox74
        '
        resources.ApplyResources(Me.CheckBox74, "CheckBox74")
        Me.CheckBox74.Name = "CheckBox74"
        Me.CheckBox74.UseVisualStyleBackColor = True
        '
        'CheckBox75
        '
        resources.ApplyResources(Me.CheckBox75, "CheckBox75")
        Me.CheckBox75.Name = "CheckBox75"
        Me.CheckBox75.UseVisualStyleBackColor = True
        '
        'GroupBox41
        '
        Me.GroupBox41.Controls.Add(Me.CheckBox76)
        Me.GroupBox41.Controls.Add(Me.CheckBox77)
        resources.ApplyResources(Me.GroupBox41, "GroupBox41")
        Me.GroupBox41.Name = "GroupBox41"
        Me.GroupBox41.TabStop = False
        '
        'CheckBox76
        '
        resources.ApplyResources(Me.CheckBox76, "CheckBox76")
        Me.CheckBox76.Name = "CheckBox76"
        Me.CheckBox76.UseVisualStyleBackColor = True
        '
        'CheckBox77
        '
        resources.ApplyResources(Me.CheckBox77, "CheckBox77")
        Me.CheckBox77.Name = "CheckBox77"
        Me.CheckBox77.UseVisualStyleBackColor = True
        '
        'GroupBox42
        '
        Me.GroupBox42.Controls.Add(Me.CheckBox78)
        Me.GroupBox42.Controls.Add(Me.CheckBox79)
        resources.ApplyResources(Me.GroupBox42, "GroupBox42")
        Me.GroupBox42.Name = "GroupBox42"
        Me.GroupBox42.TabStop = False
        '
        'CheckBox78
        '
        resources.ApplyResources(Me.CheckBox78, "CheckBox78")
        Me.CheckBox78.Name = "CheckBox78"
        Me.CheckBox78.UseVisualStyleBackColor = True
        '
        'CheckBox79
        '
        resources.ApplyResources(Me.CheckBox79, "CheckBox79")
        Me.CheckBox79.Name = "CheckBox79"
        Me.CheckBox79.UseVisualStyleBackColor = True
        '
        'GroupBox43
        '
        Me.GroupBox43.Controls.Add(Me.CheckBox80)
        Me.GroupBox43.Controls.Add(Me.CheckBox81)
        Me.GroupBox43.Controls.Add(Me.CheckBox82)
        Me.GroupBox43.Controls.Add(Me.CheckBox83)
        resources.ApplyResources(Me.GroupBox43, "GroupBox43")
        Me.GroupBox43.Name = "GroupBox43"
        Me.GroupBox43.TabStop = False
        '
        'CheckBox80
        '
        resources.ApplyResources(Me.CheckBox80, "CheckBox80")
        Me.CheckBox80.Name = "CheckBox80"
        Me.CheckBox80.UseVisualStyleBackColor = True
        '
        'CheckBox81
        '
        resources.ApplyResources(Me.CheckBox81, "CheckBox81")
        Me.CheckBox81.Name = "CheckBox81"
        Me.CheckBox81.UseVisualStyleBackColor = True
        '
        'CheckBox82
        '
        resources.ApplyResources(Me.CheckBox82, "CheckBox82")
        Me.CheckBox82.Name = "CheckBox82"
        Me.CheckBox82.UseVisualStyleBackColor = True
        '
        'CheckBox83
        '
        resources.ApplyResources(Me.CheckBox83, "CheckBox83")
        Me.CheckBox83.Name = "CheckBox83"
        Me.CheckBox83.UseVisualStyleBackColor = True
        '
        'GroupBoxEquipementsMarchands
        '
        resources.ApplyResources(Me.GroupBoxEquipementsMarchands, "GroupBoxEquipementsMarchands")
        Me.GroupBoxEquipementsMarchands.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.GroupBoxEquipementsMarchands.Controls.Add(Me.GroupBox29)
        Me.GroupBoxEquipementsMarchands.Controls.Add(Me.GroupBox30)
        Me.GroupBoxEquipementsMarchands.Controls.Add(Me.GroupBox31)
        Me.GroupBoxEquipementsMarchands.Controls.Add(Me.GroupBox32)
        Me.GroupBoxEquipementsMarchands.Controls.Add(Me.GroupBox33)
        Me.GroupBoxEquipementsMarchands.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxEquipementsMarchands.Name = "GroupBoxEquipementsMarchands"
        Me.GroupBoxEquipementsMarchands.TabStop = False
        '
        'GroupBox29
        '
        Me.GroupBox29.Controls.Add(Me.NumericUpDown7)
        Me.GroupBox29.Controls.Add(Me.Label6)
        resources.ApplyResources(Me.GroupBox29, "GroupBox29")
        Me.GroupBox29.Name = "GroupBox29"
        Me.GroupBox29.TabStop = False
        '
        'NumericUpDown7
        '
        resources.ApplyResources(Me.NumericUpDown7, "NumericUpDown7")
        Me.NumericUpDown7.Name = "NumericUpDown7"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'GroupBox30
        '
        Me.GroupBox30.Controls.Add(Me.CheckBox51)
        Me.GroupBox30.Controls.Add(Me.CheckBox52)
        Me.GroupBox30.Controls.Add(Me.CheckBox53)
        resources.ApplyResources(Me.GroupBox30, "GroupBox30")
        Me.GroupBox30.Name = "GroupBox30"
        Me.GroupBox30.TabStop = False
        '
        'CheckBox51
        '
        resources.ApplyResources(Me.CheckBox51, "CheckBox51")
        Me.CheckBox51.Name = "CheckBox51"
        Me.CheckBox51.UseVisualStyleBackColor = True
        '
        'CheckBox52
        '
        resources.ApplyResources(Me.CheckBox52, "CheckBox52")
        Me.CheckBox52.Name = "CheckBox52"
        Me.CheckBox52.UseVisualStyleBackColor = True
        '
        'CheckBox53
        '
        resources.ApplyResources(Me.CheckBox53, "CheckBox53")
        Me.CheckBox53.Name = "CheckBox53"
        Me.CheckBox53.UseVisualStyleBackColor = True
        '
        'GroupBox31
        '
        Me.GroupBox31.Controls.Add(Me.CheckBox54)
        Me.GroupBox31.Controls.Add(Me.CheckBox55)
        resources.ApplyResources(Me.GroupBox31, "GroupBox31")
        Me.GroupBox31.Name = "GroupBox31"
        Me.GroupBox31.TabStop = False
        '
        'CheckBox54
        '
        resources.ApplyResources(Me.CheckBox54, "CheckBox54")
        Me.CheckBox54.Name = "CheckBox54"
        Me.CheckBox54.UseVisualStyleBackColor = True
        '
        'CheckBox55
        '
        resources.ApplyResources(Me.CheckBox55, "CheckBox55")
        Me.CheckBox55.Name = "CheckBox55"
        Me.CheckBox55.UseVisualStyleBackColor = True
        '
        'GroupBox32
        '
        Me.GroupBox32.Controls.Add(Me.CheckBox56)
        Me.GroupBox32.Controls.Add(Me.CheckBox57)
        resources.ApplyResources(Me.GroupBox32, "GroupBox32")
        Me.GroupBox32.Name = "GroupBox32"
        Me.GroupBox32.TabStop = False
        '
        'CheckBox56
        '
        resources.ApplyResources(Me.CheckBox56, "CheckBox56")
        Me.CheckBox56.Name = "CheckBox56"
        Me.CheckBox56.UseVisualStyleBackColor = True
        '
        'CheckBox57
        '
        resources.ApplyResources(Me.CheckBox57, "CheckBox57")
        Me.CheckBox57.Name = "CheckBox57"
        Me.CheckBox57.UseVisualStyleBackColor = True
        '
        'GroupBox33
        '
        Me.GroupBox33.Controls.Add(Me.CheckBox58)
        Me.GroupBox33.Controls.Add(Me.CheckBox59)
        Me.GroupBox33.Controls.Add(Me.CheckBox60)
        Me.GroupBox33.Controls.Add(Me.CheckBox61)
        resources.ApplyResources(Me.GroupBox33, "GroupBox33")
        Me.GroupBox33.Name = "GroupBox33"
        Me.GroupBox33.TabStop = False
        '
        'CheckBox58
        '
        resources.ApplyResources(Me.CheckBox58, "CheckBox58")
        Me.CheckBox58.Name = "CheckBox58"
        Me.CheckBox58.UseVisualStyleBackColor = True
        '
        'CheckBox59
        '
        resources.ApplyResources(Me.CheckBox59, "CheckBox59")
        Me.CheckBox59.Name = "CheckBox59"
        Me.CheckBox59.UseVisualStyleBackColor = True
        '
        'CheckBox60
        '
        resources.ApplyResources(Me.CheckBox60, "CheckBox60")
        Me.CheckBox60.Name = "CheckBox60"
        Me.CheckBox60.UseVisualStyleBackColor = True
        '
        'CheckBox61
        '
        resources.ApplyResources(Me.CheckBox61, "CheckBox61")
        Me.CheckBox61.Name = "CheckBox61"
        Me.CheckBox61.UseVisualStyleBackColor = True
        '
        'GroupBoxTransport
        '
        resources.ApplyResources(Me.GroupBoxTransport, "GroupBoxTransport")
        Me.GroupBoxTransport.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBoxTransport.Controls.Add(Me.GroupBox34)
        Me.GroupBoxTransport.Controls.Add(Me.GroupBox35)
        Me.GroupBoxTransport.Controls.Add(Me.GroupBox36)
        Me.GroupBoxTransport.Controls.Add(Me.GroupBox37)
        Me.GroupBoxTransport.Controls.Add(Me.GroupBox38)
        Me.GroupBoxTransport.Cursor = System.Windows.Forms.Cursors.Hand
        Me.GroupBoxTransport.Name = "GroupBoxTransport"
        Me.GroupBoxTransport.TabStop = False
        '
        'GroupBox34
        '
        Me.GroupBox34.Controls.Add(Me.NumericUpDown8)
        Me.GroupBox34.Controls.Add(Me.Label7)
        resources.ApplyResources(Me.GroupBox34, "GroupBox34")
        Me.GroupBox34.Name = "GroupBox34"
        Me.GroupBox34.TabStop = False
        '
        'NumericUpDown8
        '
        resources.ApplyResources(Me.NumericUpDown8, "NumericUpDown8")
        Me.NumericUpDown8.Name = "NumericUpDown8"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'GroupBox35
        '
        Me.GroupBox35.Controls.Add(Me.CheckBox62)
        Me.GroupBox35.Controls.Add(Me.CheckBox63)
        Me.GroupBox35.Controls.Add(Me.CheckBox64)
        resources.ApplyResources(Me.GroupBox35, "GroupBox35")
        Me.GroupBox35.Name = "GroupBox35"
        Me.GroupBox35.TabStop = False
        '
        'CheckBox62
        '
        resources.ApplyResources(Me.CheckBox62, "CheckBox62")
        Me.CheckBox62.Name = "CheckBox62"
        Me.CheckBox62.UseVisualStyleBackColor = True
        '
        'CheckBox63
        '
        resources.ApplyResources(Me.CheckBox63, "CheckBox63")
        Me.CheckBox63.Name = "CheckBox63"
        Me.CheckBox63.UseVisualStyleBackColor = True
        '
        'CheckBox64
        '
        resources.ApplyResources(Me.CheckBox64, "CheckBox64")
        Me.CheckBox64.Name = "CheckBox64"
        Me.CheckBox64.UseVisualStyleBackColor = True
        '
        'GroupBox36
        '
        Me.GroupBox36.Controls.Add(Me.CheckBox65)
        Me.GroupBox36.Controls.Add(Me.CheckBox66)
        resources.ApplyResources(Me.GroupBox36, "GroupBox36")
        Me.GroupBox36.Name = "GroupBox36"
        Me.GroupBox36.TabStop = False
        '
        'CheckBox65
        '
        resources.ApplyResources(Me.CheckBox65, "CheckBox65")
        Me.CheckBox65.Name = "CheckBox65"
        Me.CheckBox65.UseVisualStyleBackColor = True
        '
        'CheckBox66
        '
        resources.ApplyResources(Me.CheckBox66, "CheckBox66")
        Me.CheckBox66.Name = "CheckBox66"
        Me.CheckBox66.UseVisualStyleBackColor = True
        '
        'GroupBox37
        '
        Me.GroupBox37.Controls.Add(Me.CheckBox67)
        Me.GroupBox37.Controls.Add(Me.CheckBox68)
        resources.ApplyResources(Me.GroupBox37, "GroupBox37")
        Me.GroupBox37.Name = "GroupBox37"
        Me.GroupBox37.TabStop = False
        '
        'CheckBox67
        '
        resources.ApplyResources(Me.CheckBox67, "CheckBox67")
        Me.CheckBox67.Name = "CheckBox67"
        Me.CheckBox67.UseVisualStyleBackColor = True
        '
        'CheckBox68
        '
        resources.ApplyResources(Me.CheckBox68, "CheckBox68")
        Me.CheckBox68.Name = "CheckBox68"
        Me.CheckBox68.UseVisualStyleBackColor = True
        '
        'GroupBox38
        '
        Me.GroupBox38.Controls.Add(Me.CheckBox69)
        Me.GroupBox38.Controls.Add(Me.CheckBox70)
        Me.GroupBox38.Controls.Add(Me.CheckBox71)
        Me.GroupBox38.Controls.Add(Me.CheckBox72)
        resources.ApplyResources(Me.GroupBox38, "GroupBox38")
        Me.GroupBox38.Name = "GroupBox38"
        Me.GroupBox38.TabStop = False
        '
        'CheckBox69
        '
        resources.ApplyResources(Me.CheckBox69, "CheckBox69")
        Me.CheckBox69.Name = "CheckBox69"
        Me.CheckBox69.UseVisualStyleBackColor = True
        '
        'CheckBox70
        '
        resources.ApplyResources(Me.CheckBox70, "CheckBox70")
        Me.CheckBox70.Name = "CheckBox70"
        Me.CheckBox70.UseVisualStyleBackColor = True
        '
        'CheckBox71
        '
        resources.ApplyResources(Me.CheckBox71, "CheckBox71")
        Me.CheckBox71.Name = "CheckBox71"
        Me.CheckBox71.UseVisualStyleBackColor = True
        '
        'CheckBox72
        '
        resources.ApplyResources(Me.CheckBox72, "CheckBox72")
        Me.CheckBox72.Name = "CheckBox72"
        Me.CheckBox72.UseVisualStyleBackColor = True
        '
        'GroupBoxFiltresSpecifiquesToutInfras
        '
        Me.GroupBoxFiltresSpecifiquesToutInfras.BackColor = System.Drawing.Color.Transparent
        Me.GroupBoxFiltresSpecifiquesToutInfras.Controls.Add(Me.GroupBoxSelectFiltres)
        resources.ApplyResources(Me.GroupBoxFiltresSpecifiquesToutInfras, "GroupBoxFiltresSpecifiquesToutInfras")
        Me.GroupBoxFiltresSpecifiquesToutInfras.Name = "GroupBoxFiltresSpecifiquesToutInfras"
        Me.GroupBoxFiltresSpecifiquesToutInfras.TabStop = False
        '
        'ButtonSelectionnerToutInfra
        '
        resources.ApplyResources(Me.ButtonSelectionnerToutInfra, "ButtonSelectionnerToutInfra")
        Me.ButtonSelectionnerToutInfra.Name = "ButtonSelectionnerToutInfra"
        Me.ButtonSelectionnerToutInfra.UseVisualStyleBackColor = True
        '
        'ButtonReinitialiserInfra
        '
        resources.ApplyResources(Me.ButtonReinitialiserInfra, "ButtonReinitialiserInfra")
        Me.ButtonReinitialiserInfra.Name = "ButtonReinitialiserInfra"
        Me.ButtonReinitialiserInfra.UseVisualStyleBackColor = True
        '
        'CheckBoxAdministration
        '
        resources.ApplyResources(Me.CheckBoxAdministration, "CheckBoxAdministration")
        Me.CheckBoxAdministration.Name = "CheckBoxAdministration"
        Me.CheckBoxAdministration.UseVisualStyleBackColor = True
        '
        'CheckBoxEauAssainissement
        '
        resources.ApplyResources(Me.CheckBoxEauAssainissement, "CheckBoxEauAssainissement")
        Me.CheckBoxEauAssainissement.Name = "CheckBoxEauAssainissement"
        Me.CheckBoxEauAssainissement.UseVisualStyleBackColor = True
        '
        'CheckBoxElevage
        '
        resources.ApplyResources(Me.CheckBoxElevage, "CheckBoxElevage")
        Me.CheckBoxElevage.Name = "CheckBoxElevage"
        Me.CheckBoxElevage.UseVisualStyleBackColor = True
        '
        'CheckBoxEquipementMarchand
        '
        resources.ApplyResources(Me.CheckBoxEquipementMarchand, "CheckBoxEquipementMarchand")
        Me.CheckBoxEquipementMarchand.Name = "CheckBoxEquipementMarchand"
        Me.CheckBoxEquipementMarchand.UseVisualStyleBackColor = True
        '
        'CheckBoxAgriculture
        '
        resources.ApplyResources(Me.CheckBoxAgriculture, "CheckBoxAgriculture")
        Me.CheckBoxAgriculture.Name = "CheckBoxAgriculture"
        Me.CheckBoxAgriculture.UseVisualStyleBackColor = True
        '
        'CheckBoxEnergie
        '
        resources.ApplyResources(Me.CheckBoxEnergie, "CheckBoxEnergie")
        Me.CheckBoxEnergie.Name = "CheckBoxEnergie"
        Me.CheckBoxEnergie.UseVisualStyleBackColor = True
        '
        'CheckBoxEducation
        '
        resources.ApplyResources(Me.CheckBoxEducation, "CheckBoxEducation")
        Me.CheckBoxEducation.Name = "CheckBoxEducation"
        Me.CheckBoxEducation.UseVisualStyleBackColor = True
        '
        'CheckBoxSante
        '
        resources.ApplyResources(Me.CheckBoxSante, "CheckBoxSante")
        Me.CheckBoxSante.Name = "CheckBoxSante"
        Me.CheckBoxSante.UseVisualStyleBackColor = True
        '
        'CheckBoxTransport
        '
        resources.ApplyResources(Me.CheckBoxTransport, "CheckBoxTransport")
        Me.CheckBoxTransport.Name = "CheckBoxTransport"
        Me.CheckBoxTransport.UseVisualStyleBackColor = True
        '
        'CheckBoxSportsLoisirs
        '
        resources.ApplyResources(Me.CheckBoxSportsLoisirs, "CheckBoxSportsLoisirs")
        Me.CheckBoxSportsLoisirs.Name = "CheckBoxSportsLoisirs"
        Me.CheckBoxSportsLoisirs.UseVisualStyleBackColor = True
        '
        'GroupBoxChoixInfras
        '
        Me.GroupBoxChoixInfras.Controls.Add(Me.LabelSourceEducation)
        Me.GroupBoxChoixInfras.Controls.Add(Me.ButtonAfficherResultatSansFilter)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxSportsLoisirs)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxTransport)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxSante)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxEducation)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxEnergie)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxAgriculture)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxEquipementMarchand)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxElevage)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxEauAssainissement)
        Me.GroupBoxChoixInfras.Controls.Add(Me.CheckBoxAdministration)
        Me.GroupBoxChoixInfras.Controls.Add(Me.ButtonReinitialiserInfra)
        Me.GroupBoxChoixInfras.Controls.Add(Me.ButtonSelectionnerToutInfra)
        resources.ApplyResources(Me.GroupBoxChoixInfras, "GroupBoxChoixInfras")
        Me.GroupBoxChoixInfras.Name = "GroupBoxChoixInfras"
        Me.GroupBoxChoixInfras.TabStop = False
        '
        'LabelSourceEducation
        '
        resources.ApplyResources(Me.LabelSourceEducation, "LabelSourceEducation")
        Me.LabelSourceEducation.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.LabelSourceEducation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelSourceEducation.Name = "LabelSourceEducation"
        '
        'ButtonAfficherResultatSansFilter
        '
        resources.ApplyResources(Me.ButtonAfficherResultatSansFilter, "ButtonAfficherResultatSansFilter")
        Me.ButtonAfficherResultatSansFilter.Name = "ButtonAfficherResultatSansFilter"
        Me.ButtonAfficherResultatSansFilter.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        resources.ApplyResources(Me.DataGridView1, "DataGridView1")
        Me.DataGridView1.Name = "DataGridView1"
        '
        'OpenFileDialogEducation
        '
        Me.OpenFileDialogEducation.FileName = "OpenFileDialogEducation"
        '
        'DataSetExcelko1
        '
        Me.DataSetExcelko1.DataSetName = "DataSetExcelko1"
        Me.DataSetExcelko1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ButtonVoirRapportDetaille
        '
        resources.ApplyResources(Me.ButtonVoirRapportDetaille, "ButtonVoirRapportDetaille")
        Me.ButtonVoirRapportDetaille.Name = "ButtonVoirRapportDetaille"
        Me.ButtonVoirRapportDetaille.UseVisualStyleBackColor = True
        '
        'Form1
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange
        Me.Controls.Add(Me.ButtonVoirRapportDetaille)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.GroupBoxSportsEtLoisirs)
        Me.Controls.Add(Me.GroupBoxEnergie)
        Me.Controls.Add(Me.GroupBoxEducation)
        Me.Controls.Add(Me.GroupBoxEauEtAssainissement)
        Me.Controls.Add(Me.GroupBoxTransport)
        Me.Controls.Add(Me.GroupBoxSante)
        Me.Controls.Add(Me.GroupBoxEquipementsMarchands)
        Me.Controls.Add(Me.GroupBoxFiltresGeographiques)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.GroupBoxChoixInfras)
        Me.Controls.Add(Me.GroupBoxFiltresSpecifiquesToutInfras)
        Me.Controls.Add(Me.GroupBoxAgricultureEtElevage)
        Me.Controls.Add(Me.GroupBoxAdministration)
        Me.HelpButton = True
        Me.Name = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBoxFiltresGeographiques.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBoxSelectFiltres.ResumeLayout(False)
        Me.GroupBoxEducation.ResumeLayout(False)
        Me.GroupBoxNombreDEleve.ResumeLayout(False)
        Me.GroupBoxNombreDEleve.PerformLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxNiveauEnseigne.ResumeLayout(False)
        Me.GroupBoxNiveauEnseigne.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBoxEtatInfraEducation.ResumeLayout(False)
        Me.GroupBoxEtatInfraEducation.PerformLayout()
        Me.GroupBoxAdministration.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        Me.GroupBoxSante.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxDiplomeDuReferent.ResumeLayout(False)
        Me.GroupBoxDiplomeDuReferent.PerformLayout()
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox12.PerformLayout()
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox13.PerformLayout()
        Me.GroupBox14.ResumeLayout(False)
        Me.GroupBox14.PerformLayout()
        Me.GroupBoxEauEtAssainissement.ResumeLayout(False)
        Me.GroupBox15.ResumeLayout(False)
        Me.GroupBox15.PerformLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxTypeDInstallation.ResumeLayout(False)
        Me.GroupBoxTypeDInstallation.PerformLayout()
        Me.GroupBox17.ResumeLayout(False)
        Me.GroupBox17.PerformLayout()
        Me.GroupBox18.ResumeLayout(False)
        Me.GroupBox18.PerformLayout()
        Me.GroupBox19.ResumeLayout(False)
        Me.GroupBox19.PerformLayout()
        Me.GroupBoxSportsEtLoisirs.ResumeLayout(False)
        Me.GroupBox16.ResumeLayout(False)
        Me.GroupBox16.PerformLayout()
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox20.ResumeLayout(False)
        Me.GroupBox20.PerformLayout()
        Me.GroupBox21.ResumeLayout(False)
        Me.GroupBox21.PerformLayout()
        Me.GroupBox22.ResumeLayout(False)
        Me.GroupBox22.PerformLayout()
        Me.GroupBox23.ResumeLayout(False)
        Me.GroupBox23.PerformLayout()
        Me.GroupBoxAgricultureEtElevage.ResumeLayout(False)
        Me.GroupBox24.ResumeLayout(False)
        Me.GroupBox24.PerformLayout()
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox25.ResumeLayout(False)
        Me.GroupBox25.PerformLayout()
        Me.GroupBox26.ResumeLayout(False)
        Me.GroupBox26.PerformLayout()
        Me.GroupBox27.ResumeLayout(False)
        Me.GroupBox27.PerformLayout()
        Me.GroupBox28.ResumeLayout(False)
        Me.GroupBox28.PerformLayout()
        Me.GroupBoxEnergie.ResumeLayout(False)
        Me.GroupBox39.ResumeLayout(False)
        Me.GroupBox39.PerformLayout()
        CType(Me.NumericUpDown9, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox40.ResumeLayout(False)
        Me.GroupBox40.PerformLayout()
        Me.GroupBox41.ResumeLayout(False)
        Me.GroupBox41.PerformLayout()
        Me.GroupBox42.ResumeLayout(False)
        Me.GroupBox42.PerformLayout()
        Me.GroupBox43.ResumeLayout(False)
        Me.GroupBox43.PerformLayout()
        Me.GroupBoxEquipementsMarchands.ResumeLayout(False)
        Me.GroupBox29.ResumeLayout(False)
        Me.GroupBox29.PerformLayout()
        CType(Me.NumericUpDown7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox30.ResumeLayout(False)
        Me.GroupBox30.PerformLayout()
        Me.GroupBox31.ResumeLayout(False)
        Me.GroupBox31.PerformLayout()
        Me.GroupBox32.ResumeLayout(False)
        Me.GroupBox32.PerformLayout()
        Me.GroupBox33.ResumeLayout(False)
        Me.GroupBox33.PerformLayout()
        Me.GroupBoxTransport.ResumeLayout(False)
        Me.GroupBox34.ResumeLayout(False)
        Me.GroupBox34.PerformLayout()
        CType(Me.NumericUpDown8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox35.ResumeLayout(False)
        Me.GroupBox35.PerformLayout()
        Me.GroupBox36.ResumeLayout(False)
        Me.GroupBox36.PerformLayout()
        Me.GroupBox37.ResumeLayout(False)
        Me.GroupBox37.PerformLayout()
        Me.GroupBox38.ResumeLayout(False)
        Me.GroupBox38.PerformLayout()
        Me.GroupBoxFiltresSpecifiquesToutInfras.ResumeLayout(False)
        Me.GroupBoxChoixInfras.ResumeLayout(False)
        Me.GroupBoxChoixInfras.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSetExcelko1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TreeViewFiltresGeographiques As TreeView
    Friend WithEvents GroupBoxFiltresGeographiques As GroupBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImporterSourceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EauEtAssainissementToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SportsEtLoisirsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AdministrationToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SantéToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EducationToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AgricultureToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ElevageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TransportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EquipementsMarchandsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EnergieToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AProposDeLoutilToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents QuitterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GroupBoxSelectFiltres As GroupBox
    Friend WithEvents ButtonAppliquerFiltres As Button
    Friend WithEvents ComboBoxInfraAFiltrer As ComboBox
    Friend WithEvents GroupBoxEducation As GroupBox
    Friend WithEvents GroupBoxNombreDEleve As GroupBox
    Friend WithEvents NumericUpDown1 As NumericUpDown
    Friend WithEvents LabelNombreEleveSuperieurA As Label
    Friend WithEvents GroupBoxAdministration As GroupBox
    Friend WithEvents GroupBoxSante As GroupBox
    Friend WithEvents GroupBoxEauEtAssainissement As GroupBox
    Friend WithEvents GroupBoxSportsEtLoisirs As GroupBox
    Friend WithEvents GroupBoxAgricultureEtElevage As GroupBox
    Friend WithEvents GroupBox24 As GroupBox
    Friend WithEvents NumericUpDown6 As NumericUpDown
    Friend WithEvents Label5 As Label
    Friend WithEvents GroupBox25 As GroupBox
    Friend WithEvents CheckBox40 As CheckBox
    Friend WithEvents CheckBox41 As CheckBox
    Friend WithEvents CheckBox42 As CheckBox
    Friend WithEvents GroupBox26 As GroupBox
    Friend WithEvents CheckBox43 As CheckBox
    Friend WithEvents CheckBox44 As CheckBox
    Friend WithEvents GroupBox27 As GroupBox
    Friend WithEvents CheckBox45 As CheckBox
    Friend WithEvents CheckBox46 As CheckBox
    Friend WithEvents GroupBox28 As GroupBox
    Friend WithEvents CheckBox47 As CheckBox
    Friend WithEvents CheckBox48 As CheckBox
    Friend WithEvents CheckBox49 As CheckBox
    Friend WithEvents CheckBox50 As CheckBox
    Friend WithEvents GroupBoxEnergie As GroupBox
    Friend WithEvents GroupBox39 As GroupBox
    Friend WithEvents NumericUpDown9 As NumericUpDown
    Friend WithEvents Label8 As Label
    Friend WithEvents GroupBox40 As GroupBox
    Friend WithEvents CheckBox73 As CheckBox
    Friend WithEvents CheckBox74 As CheckBox
    Friend WithEvents CheckBox75 As CheckBox
    Friend WithEvents GroupBox41 As GroupBox
    Friend WithEvents CheckBox76 As CheckBox
    Friend WithEvents CheckBox77 As CheckBox
    Friend WithEvents GroupBox42 As GroupBox
    Friend WithEvents CheckBox78 As CheckBox
    Friend WithEvents CheckBox79 As CheckBox
    Friend WithEvents GroupBox43 As GroupBox
    Friend WithEvents CheckBox80 As CheckBox
    Friend WithEvents CheckBox81 As CheckBox
    Friend WithEvents CheckBox82 As CheckBox
    Friend WithEvents CheckBox83 As CheckBox
    Friend WithEvents GroupBox16 As GroupBox
    Friend WithEvents NumericUpDown5 As NumericUpDown
    Friend WithEvents Label4 As Label
    Friend WithEvents GroupBox20 As GroupBox
    Friend WithEvents CheckBox15 As CheckBox
    Friend WithEvents CheckBox16 As CheckBox
    Friend WithEvents CheckBox17 As CheckBox
    Friend WithEvents GroupBox21 As GroupBox
    Friend WithEvents CheckBox32 As CheckBox
    Friend WithEvents CheckBox33 As CheckBox
    Friend WithEvents GroupBox22 As GroupBox
    Friend WithEvents CheckBox34 As CheckBox
    Friend WithEvents CheckBox35 As CheckBox
    Friend WithEvents GroupBox23 As GroupBox
    Friend WithEvents CheckBox36 As CheckBox
    Friend WithEvents CheckBox37 As CheckBox
    Friend WithEvents CheckBox38 As CheckBox
    Friend WithEvents CheckBox39 As CheckBox
    Friend WithEvents GroupBox15 As GroupBox
    Friend WithEvents NumericUpDown4 As NumericUpDown
    Friend WithEvents Label3 As Label
    Friend WithEvents GroupBoxTypeDInstallation As GroupBox
    Friend WithEvents CheckBoxChâteauDEau As CheckBox
    Friend WithEvents CheckBoxPuits As CheckBox
    Friend WithEvents CheckBoxBorneFontaine As CheckBox
    Friend WithEvents GroupBox17 As GroupBox
    Friend WithEvents CheckBox18 As CheckBox
    Friend WithEvents CheckBox25 As CheckBox
    Friend WithEvents GroupBox18 As GroupBox
    Friend WithEvents CheckBox26 As CheckBox
    Friend WithEvents CheckBox27 As CheckBox
    Friend WithEvents GroupBox19 As GroupBox
    Friend WithEvents CheckBox28 As CheckBox
    Friend WithEvents CheckBox29 As CheckBox
    Friend WithEvents CheckBox30 As CheckBox
    Friend WithEvents CheckBox31 As CheckBox
    Friend WithEvents GroupBox11 As GroupBox
    Friend WithEvents NumericUpDown3 As NumericUpDown
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBoxDiplomeDuReferent As GroupBox
    Friend WithEvents CheckBoxInfirmier As CheckBox
    Friend WithEvents CheckBoxSageFemme As CheckBox
    Friend WithEvents CheckBoxDocteur As CheckBox
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents CheckBoxCSBII As CheckBox
    Friend WithEvents CheckBoxCSBI As CheckBox
    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents CheckBox19 As CheckBox
    Friend WithEvents CheckBox20 As CheckBox
    Friend WithEvents GroupBox14 As GroupBox
    Friend WithEvents CheckBox21 As CheckBox
    Friend WithEvents CheckBox22 As CheckBox
    Friend WithEvents CheckBox23 As CheckBox
    Friend WithEvents CheckBox24 As CheckBox
    Friend WithEvents GroupBox6 As GroupBox
    Friend WithEvents NumericUpDown2 As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox8 As GroupBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents GroupBox9 As GroupBox
    Friend WithEvents CheckBox6 As CheckBox
    Friend WithEvents CheckBox7 As CheckBox
    Friend WithEvents GroupBox10 As GroupBox
    Friend WithEvents CheckBox8 As CheckBox
    Friend WithEvents CheckBox9 As CheckBox
    Friend WithEvents CheckBox10 As CheckBox
    Friend WithEvents CheckBox14 As CheckBox
    Friend WithEvents GroupBoxNiveauEnseigne As GroupBox
    Friend WithEvents CheckBoxLycee As CheckBox
    Friend WithEvents CheckBoxCollege As CheckBox
    Friend WithEvents CheckBoxPrimaire As CheckBox
    Friend WithEvents CheckBoxPrescolaire As CheckBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents CheckBoxEcolePublique As CheckBox
    Friend WithEvents CheckBoxEcolePrive As CheckBox
    Friend WithEvents GroupBoxEtatInfraEducation As GroupBox
    Friend WithEvents CheckBoxInutilisableEducation As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents CheckBoxEtatUsageEducation As CheckBox
    Friend WithEvents CheckBoxBonEtatEducation As CheckBox
    Friend WithEvents GroupBoxEquipementsMarchands As GroupBox
    Friend WithEvents GroupBox29 As GroupBox
    Friend WithEvents NumericUpDown7 As NumericUpDown
    Friend WithEvents Label6 As Label
    Friend WithEvents GroupBox30 As GroupBox
    Friend WithEvents CheckBox51 As CheckBox
    Friend WithEvents CheckBox52 As CheckBox
    Friend WithEvents CheckBox53 As CheckBox
    Friend WithEvents GroupBox31 As GroupBox
    Friend WithEvents CheckBox54 As CheckBox
    Friend WithEvents CheckBox55 As CheckBox
    Friend WithEvents GroupBox32 As GroupBox
    Friend WithEvents CheckBox56 As CheckBox
    Friend WithEvents CheckBox57 As CheckBox
    Friend WithEvents GroupBox33 As GroupBox
    Friend WithEvents CheckBox58 As CheckBox
    Friend WithEvents CheckBox59 As CheckBox
    Friend WithEvents CheckBox60 As CheckBox
    Friend WithEvents CheckBox61 As CheckBox
    Friend WithEvents GroupBoxTransport As GroupBox
    Friend WithEvents GroupBox34 As GroupBox
    Friend WithEvents NumericUpDown8 As NumericUpDown
    Friend WithEvents Label7 As Label
    Friend WithEvents GroupBox35 As GroupBox
    Friend WithEvents CheckBox62 As CheckBox
    Friend WithEvents CheckBox63 As CheckBox
    Friend WithEvents CheckBox64 As CheckBox
    Friend WithEvents GroupBox36 As GroupBox
    Friend WithEvents CheckBox65 As CheckBox
    Friend WithEvents CheckBox66 As CheckBox
    Friend WithEvents GroupBox37 As GroupBox
    Friend WithEvents CheckBox67 As CheckBox
    Friend WithEvents CheckBox68 As CheckBox
    Friend WithEvents GroupBox38 As GroupBox
    Friend WithEvents CheckBox69 As CheckBox
    Friend WithEvents CheckBox70 As CheckBox
    Friend WithEvents CheckBox71 As CheckBox
    Friend WithEvents CheckBox72 As CheckBox
    Friend WithEvents GroupBoxFiltresSpecifiquesToutInfras As GroupBox
    Friend WithEvents ButtonCancelFiltres As Button
    Friend WithEvents ButtonSelectAllFiltres As Button
    Friend WithEvents ButtonSelectionnerToutInfra As Button
    Friend WithEvents ButtonReinitialiserInfra As Button
    Friend WithEvents CheckBoxAdministration As CheckBox
    Friend WithEvents CheckBoxEauAssainissement As CheckBox
    Friend WithEvents CheckBoxElevage As CheckBox
    Friend WithEvents CheckBoxEquipementMarchand As CheckBox
    Friend WithEvents CheckBoxAgriculture As CheckBox
    Friend WithEvents CheckBoxEnergie As CheckBox
    Friend WithEvents CheckBoxEducation As CheckBox
    Friend WithEvents CheckBoxSante As CheckBox
    Friend WithEvents CheckBoxTransport As CheckBox
    Friend WithEvents CheckBoxSportsLoisirs As CheckBox
    Friend WithEvents GroupBoxChoixInfras As GroupBox
    Friend WithEvents ButtonAfficherResultatSansFilter As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents OpenFileDialogEducation As OpenFileDialog
    Friend WithEvents LabelSourceEducation As Label
    Friend WithEvents DataSetExcelko1 As DataSetExcelko1
    Friend WithEvents ButtonVoirRapportDetaille As Button
End Class
