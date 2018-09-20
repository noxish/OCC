<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TestDataSet3 = New OutlookAddIn2.TestDataSet3()
        Me.ContactsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ContactsTableAdapter = New OutlookAddIn2.TestDataSet3TableAdapters.ContactsTableAdapter()
        Me.SalutationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FirstnameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TestDataSet3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ContactsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SalutationDataGridViewTextBoxColumn, Me.NameDataGridViewTextBoxColumn, Me.FirstnameDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.ContactsBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(388, 221)
        Me.DataGridView1.TabIndex = 0
        '
        'TestDataSet3
        '
        Me.TestDataSet3.DataSetName = "TestDataSet3"
        Me.TestDataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ContactsBindingSource
        '
        Me.ContactsBindingSource.DataMember = "Contacts"
        Me.ContactsBindingSource.DataSource = Me.TestDataSet3
        '
        'ContactsTableAdapter
        '
        Me.ContactsTableAdapter.ClearBeforeFill = True
        '
        'SalutationDataGridViewTextBoxColumn
        '
        Me.SalutationDataGridViewTextBoxColumn.DataPropertyName = "Salutation"
        Me.SalutationDataGridViewTextBoxColumn.HeaderText = "Salutation"
        Me.SalutationDataGridViewTextBoxColumn.Name = "SalutationDataGridViewTextBoxColumn"
        Me.SalutationDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NameDataGridViewTextBoxColumn
        '
        Me.NameDataGridViewTextBoxColumn.DataPropertyName = "Name"
        Me.NameDataGridViewTextBoxColumn.HeaderText = "Name"
        Me.NameDataGridViewTextBoxColumn.Name = "NameDataGridViewTextBoxColumn"
        Me.NameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FirstnameDataGridViewTextBoxColumn
        '
        Me.FirstnameDataGridViewTextBoxColumn.DataPropertyName = "Firstname"
        Me.FirstnameDataGridViewTextBoxColumn.HeaderText = "Firstname"
        Me.FirstnameDataGridViewTextBoxColumn.Name = "FirstnameDataGridViewTextBoxColumn"
        Me.FirstnameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Button1.Location = New System.Drawing.Point(12, 239)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 31)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(412, 282)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form1"
        Me.ShowIcon = False
        Me.Text = "Kontakt Tool"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TestDataSet3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ContactsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents TestDataSet3 As TestDataSet3
    Friend WithEvents ContactsBindingSource As Windows.Forms.BindingSource
    Friend WithEvents ContactsTableAdapter As TestDataSet3TableAdapters.ContactsTableAdapter
    Friend WithEvents SalutationDataGridViewTextBoxColumn As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameDataGridViewTextBoxColumn As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FirstnameDataGridViewTextBoxColumn As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
