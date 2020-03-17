<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class taskbar_button
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OpenFileLocationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PinToTaskbarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UnpinFromTaskbarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(5, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(44, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Highlight
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(5, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 5)
        Me.Label1.TabIndex = 1
        Me.Label1.UseCompatibleTextRendering = True
        Me.Label1.Visible = False
        '
        'Menu
        '
        Me.Menu.BackColor = System.Drawing.Color.FromArgb(CType(CType(30, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(30, Byte), Integer))
        Me.Menu.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenFileLocationToolStripMenuItem, Me.PinToTaskbarToolStripMenuItem, Me.UnpinFromTaskbarToolStripMenuItem, Me.CloseToolStripMenuItem})
        Me.Menu.Margin = New System.Windows.Forms.Padding(2)
        Me.Menu.MinimumSize = New System.Drawing.Size(250, 0)
        Me.Menu.Name = "Menu"
        Me.Menu.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.Menu.Size = New System.Drawing.Size(250, 168)
        '
        'OpenFileLocationToolStripMenuItem
        '
        Me.OpenFileLocationToolStripMenuItem.AutoSize = False
        Me.OpenFileLocationToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.OpenFileLocationToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.OpenFileLocationToolStripMenuItem.Margin = New System.Windows.Forms.Padding(0, 10, 0, 0)
        Me.OpenFileLocationToolStripMenuItem.Name = "OpenFileLocationToolStripMenuItem"
        Me.OpenFileLocationToolStripMenuItem.Size = New System.Drawing.Size(250, 33)
        Me.OpenFileLocationToolStripMenuItem.Text = "Open file location"
        '
        'PinToTaskbarToolStripMenuItem
        '
        Me.PinToTaskbarToolStripMenuItem.AutoSize = False
        Me.PinToTaskbarToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.PinToTaskbarToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.PinToTaskbarToolStripMenuItem.Name = "PinToTaskbarToolStripMenuItem"
        Me.PinToTaskbarToolStripMenuItem.Size = New System.Drawing.Size(250, 33)
        Me.PinToTaskbarToolStripMenuItem.Text = "Pin to taskbar"
        '
        'UnpinFromTaskbarToolStripMenuItem
        '
        Me.UnpinFromTaskbarToolStripMenuItem.AutoSize = False
        Me.UnpinFromTaskbarToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.UnpinFromTaskbarToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.UnpinFromTaskbarToolStripMenuItem.Name = "UnpinFromTaskbarToolStripMenuItem"
        Me.UnpinFromTaskbarToolStripMenuItem.Size = New System.Drawing.Size(250, 33)
        Me.UnpinFromTaskbarToolStripMenuItem.Text = "Unpin from taskbar"
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.AutoSize = False
        Me.CloseToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(250, 33)
        Me.CloseToolStripMenuItem.Text = "Close"
        '
        'taskbar_button
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ContextMenuStrip = Me.Menu
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "taskbar_button"
        Me.Size = New System.Drawing.Size(44, 40)
        Me.Menu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Menu As ContextMenuStrip
    Friend WithEvents OpenFileLocationToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PinToTaskbarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UnpinFromTaskbarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CloseToolStripMenuItem As ToolStripMenuItem
End Class
