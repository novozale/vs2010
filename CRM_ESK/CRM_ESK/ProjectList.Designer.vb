<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProjectList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.Button24 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.DateTimePicker3 = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DateTimePicker4 = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.SfDataGrid4 = New Syncfusion.WinForms.DataGrid.SfDataGrid()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.GroupBox9.SuspendLayout()
        CType(Me.SfDataGrid4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.Button24)
        Me.GroupBox9.Location = New System.Drawing.Point(427, 4)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(152, 45)
        Me.GroupBox9.TabIndex = 55
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Дополнительно"
        '
        'Button24
        '
        Me.Button24.Location = New System.Drawing.Point(6, 15)
        Me.Button24.Name = "Button24"
        Me.Button24.Size = New System.Drawing.Size(138, 24)
        Me.Button24.TabIndex = 0
        Me.Button24.Text = "Включить группировку"
        Me.Button24.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Location = New System.Drawing.Point(802, 20)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(92, 23)
        Me.Button17.TabIndex = 54
        Me.Button17.Text = "Обновить"
        Me.Button17.UseVisualStyleBackColor = True
        '
        'DateTimePicker3
        '
        Me.DateTimePicker3.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker3.Location = New System.Drawing.Point(315, 22)
        Me.DateTimePicker3.Name = "DateTimePicker3"
        Me.DateTimePicker3.Size = New System.Drawing.Size(106, 20)
        Me.DateTimePicker3.TabIndex = 53
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label11.Location = New System.Drawing.Point(280, 23)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(36, 21)
        Me.Label11.TabIndex = 52
        Me.Label11.Text = "По"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label12.Location = New System.Drawing.Point(136, 21)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(18, 21)
        Me.Label12.TabIndex = 51
        Me.Label12.Text = "С"
        '
        'DateTimePicker4
        '
        Me.DateTimePicker4.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker4.Location = New System.Drawing.Point(163, 21)
        Me.DateTimePicker4.Name = "DateTimePicker4"
        Me.DateTimePicker4.Size = New System.Drawing.Size(107, 20)
        Me.DateTimePicker4.TabIndex = 50
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label13.Location = New System.Drawing.Point(6, 20)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(132, 21)
        Me.Label13.TabIndex = 49
        Me.Label13.Text = "Список Проектов"
        '
        'SfDataGrid4
        '
        Me.SfDataGrid4.AccessibleName = "Table"
        Me.SfDataGrid4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SfDataGrid4.CanOverrideStyle = True
        Me.SfDataGrid4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.SfDataGrid4.Location = New System.Drawing.Point(8, 55)
        Me.SfDataGrid4.Name = "SfDataGrid4"
        Me.SfDataGrid4.Size = New System.Drawing.Size(1560, 656)
        Me.SfDataGrid4.Style.CaptionSummaryRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.CellStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.FilterRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.GroupDropAreaItemStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.GroupDropAreaStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.GroupSummaryRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.HeaderStyle.FilterIconColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer))
        Me.SfDataGrid4.Style.HeaderStyle.FilterIconSize = New System.Drawing.Size(16, 16)
        Me.SfDataGrid4.Style.HeaderStyle.Font.Bold = True
        Me.SfDataGrid4.Style.HeaderStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.HeaderStyle.SortIcon = Nothing
        Me.SfDataGrid4.Style.HeaderStyle.SortIconColor = System.Drawing.Color.Blue
        Me.SfDataGrid4.Style.IndentCellStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.PreviewRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.RowHeaderStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.StackedHeaderStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.TableSummaryRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.ToolTipStyle.Font = New System.Drawing.Font("Arial", 8.5!)
        Me.SfDataGrid4.Style.UnboundRowStyle.Font.Facename = "Arial"
        Me.SfDataGrid4.Style.ValidationErrorToolTipStyle.Font = New System.Drawing.Font("Arial", 8.5!)
        Me.SfDataGrid4.TabIndex = 56
        Me.SfDataGrid4.Text = "SfDataGrid4"
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button5.Location = New System.Drawing.Point(1399, 717)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(170, 30)
        Me.Button5.TabIndex = 58
        Me.Button5.Text = "Отмена"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button4.Location = New System.Drawing.Point(1221, 717)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(170, 30)
        Me.Button4.TabIndex = 57
        Me.Button4.Text = "Выбрать проект"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.TextBox3)
        Me.GroupBox7.Location = New System.Drawing.Point(585, 4)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(200, 46)
        Me.GroupBox7.TabIndex = 59
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Быстрый поиск"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(11, 18)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(182, 20)
        Me.TextBox3.TabIndex = 0
        '
        'ProjectList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1575, 750)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.SfDataGrid4)
        Me.Controls.Add(Me.GroupBox9)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.DateTimePicker3)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.DateTimePicker4)
        Me.Controls.Add(Me.Label13)
        Me.KeyPreview = True
        Me.Name = "ProjectList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Список проектов"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox9.ResumeLayout(False)
        CType(Me.SfDataGrid4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents Button24 As System.Windows.Forms.Button
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker3 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents SfDataGrid4 As Syncfusion.WinForms.DataGrid.SfDataGrid
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
End Class
