<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 上傳考勤
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(上傳考勤))
        Me.Bu讀取上傳 = New System.Windows.Forms.Button
        Me.DGV1 = New System.Windows.Forms.DataGridView
        Me.DGV2 = New System.Windows.Forms.DataGridView
        Me.清空資料庫 = New System.Windows.Forms.Button
        Me.資料回填 = New System.Windows.Forms.Button
        Me.時間 = New System.Windows.Forms.Label
        Me.La1筆數 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.改公司 = New System.Windows.Forms.Button
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Bu讀取上傳
        '
        Me.Bu讀取上傳.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Bu讀取上傳.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(158, Byte), Integer))
        Me.Bu讀取上傳.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Bu讀取上傳.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.Bu讀取上傳.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Bu讀取上傳.Font = New System.Drawing.Font("微軟正黑體", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Bu讀取上傳.ForeColor = System.Drawing.Color.Black
        Me.Bu讀取上傳.Location = New System.Drawing.Point(35, 32)
        Me.Bu讀取上傳.Name = "Bu讀取上傳"
        Me.Bu讀取上傳.Size = New System.Drawing.Size(80, 42)
        Me.Bu讀取上傳.TabIndex = 0
        Me.Bu讀取上傳.Text = "上傳"
        Me.Bu讀取上傳.UseVisualStyleBackColor = False
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToDeleteRows = False
        Me.DGV1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Location = New System.Drawing.Point(26, 93)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.ReadOnly = True
        Me.DGV1.RowTemplate.Height = 24
        Me.DGV1.Size = New System.Drawing.Size(353, 396)
        Me.DGV1.TabIndex = 1
        '
        'DGV2
        '
        Me.DGV2.AllowUserToAddRows = False
        Me.DGV2.AllowUserToDeleteRows = False
        Me.DGV2.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(243, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.DGV2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV2.GridColor = System.Drawing.Color.Silver
        Me.DGV2.Location = New System.Drawing.Point(26, 93)
        Me.DGV2.Name = "DGV2"
        Me.DGV2.ReadOnly = True
        Me.DGV2.RowTemplate.Height = 24
        Me.DGV2.Size = New System.Drawing.Size(1027, 487)
        Me.DGV2.TabIndex = 3
        '
        '清空資料庫
        '
        Me.清空資料庫.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.清空資料庫.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(158, Byte), Integer))
        Me.清空資料庫.Cursor = System.Windows.Forms.Cursors.Hand
        Me.清空資料庫.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.清空資料庫.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.清空資料庫.Font = New System.Drawing.Font("微軟正黑體", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.清空資料庫.ForeColor = System.Drawing.Color.Black
        Me.清空資料庫.Location = New System.Drawing.Point(209, 32)
        Me.清空資料庫.Name = "清空資料庫"
        Me.清空資料庫.Size = New System.Drawing.Size(82, 42)
        Me.清空資料庫.TabIndex = 5
        Me.清空資料庫.Text = "清空"
        Me.清空資料庫.UseVisualStyleBackColor = False
        '
        '資料回填
        '
        Me.資料回填.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.資料回填.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(158, Byte), Integer))
        Me.資料回填.Cursor = System.Windows.Forms.Cursors.Hand
        Me.資料回填.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.資料回填.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.資料回填.Font = New System.Drawing.Font("微軟正黑體", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.資料回填.ForeColor = System.Drawing.Color.Black
        Me.資料回填.Location = New System.Drawing.Point(121, 32)
        Me.資料回填.Name = "資料回填"
        Me.資料回填.Size = New System.Drawing.Size(82, 42)
        Me.資料回填.TabIndex = 6
        Me.資料回填.Text = "回填"
        Me.資料回填.UseVisualStyleBackColor = False
        '
        '時間
        '
        Me.時間.AutoSize = True
        Me.時間.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.時間.ForeColor = System.Drawing.Color.Black
        Me.時間.Location = New System.Drawing.Point(23, 599)
        Me.時間.Name = "時間"
        Me.時間.Size = New System.Drawing.Size(60, 17)
        Me.時間.TabIndex = 8
        Me.時間.Text = "版本時間"
        '
        'La1筆數
        '
        Me.La1筆數.AutoSize = True
        Me.La1筆數.Font = New System.Drawing.Font("微軟正黑體", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.La1筆數.ForeColor = System.Drawing.Color.Black
        Me.La1筆數.Location = New System.Drawing.Point(307, 41)
        Me.La1筆數.Name = "La1筆數"
        Me.La1筆數.Size = New System.Drawing.Size(84, 25)
        Me.La1筆數.TabIndex = 9
        Me.La1筆數.Text = "筆數：0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微軟正黑體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(879, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(174, 24)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "(資料已填入資料庫)"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(995, 599)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "ver.02.03"
        '
        '改公司
        '
        Me.改公司.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.改公司.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(158, Byte), Integer))
        Me.改公司.Cursor = System.Windows.Forms.Cursors.Hand
        Me.改公司.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.改公司.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.改公司.Font = New System.Drawing.Font("微軟正黑體", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.改公司.ForeColor = System.Drawing.Color.Black
        Me.改公司.Location = New System.Drawing.Point(297, 32)
        Me.改公司.Name = "改公司"
        Me.改公司.Size = New System.Drawing.Size(82, 42)
        Me.改公司.TabIndex = 12
        Me.改公司.Text = "改公司"
        Me.改公司.UseVisualStyleBackColor = False
        Me.改公司.Visible = False
        '
        '上傳考勤
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(174, Byte), Integer), CType(CType(217, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1094, 636)
        Me.Controls.Add(Me.改公司)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.La1筆數)
        Me.Controls.Add(Me.時間)
        Me.Controls.Add(Me.DGV2)
        Me.Controls.Add(Me.資料回填)
        Me.Controls.Add(Me.清空資料庫)
        Me.Controls.Add(Me.DGV1)
        Me.Controls.Add(Me.Bu讀取上傳)
        Me.Font = New System.Drawing.Font("新細明體", 10.0!)
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "上傳考勤"
        Me.Text = "上傳考勤"
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Bu讀取上傳 As System.Windows.Forms.Button
    Friend WithEvents DGV1 As System.Windows.Forms.DataGridView
    Friend WithEvents DGV2 As System.Windows.Forms.DataGridView
    Friend WithEvents 清空資料庫 As System.Windows.Forms.Button
    Friend WithEvents 資料回填 As System.Windows.Forms.Button
    Friend WithEvents 時間 As System.Windows.Forms.Label
    Friend WithEvents La1筆數 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents 改公司 As System.Windows.Forms.Button
End Class
