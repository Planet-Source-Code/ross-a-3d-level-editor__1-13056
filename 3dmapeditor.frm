VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "3dmapeditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Index           =   0
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   7155
      TabIndex        =   10
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   225
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   235
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   224
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   234
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   223
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   233
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   222
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   232
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   221
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   231
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   220
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   230
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   219
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   229
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   218
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   228
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   217
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   227
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   216
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   226
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   215
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   225
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   214
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   224
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   213
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   223
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   212
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   222
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   211
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   221
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   210
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   220
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   209
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   219
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   208
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   218
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   207
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   217
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   206
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   216
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   205
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   215
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   204
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   214
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   203
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   213
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   202
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   212
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   201
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   211
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   200
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   210
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   199
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   209
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   198
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   208
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   197
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   207
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   196
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   206
         Top             =   6240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   195
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   205
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   194
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   204
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   193
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   203
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   192
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   202
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   191
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   201
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   190
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   200
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   189
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   199
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   188
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   198
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   187
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   197
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   186
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   196
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   185
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   195
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   184
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   194
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   183
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   193
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   182
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   192
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   181
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   191
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   180
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   190
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   179
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   189
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   178
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   188
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   177
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   187
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   176
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   186
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   175
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   185
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   174
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   184
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   173
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   183
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   172
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   182
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   171
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   181
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   170
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   180
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   169
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   179
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   168
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   178
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   167
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   177
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   166
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   176
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   165
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   175
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   164
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   174
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   163
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   173
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   162
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   172
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   161
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   171
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   160
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   170
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   159
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   169
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   158
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   168
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   157
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   167
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   156
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   166
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   155
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   165
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   154
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   164
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   153
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   163
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   152
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   162
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   151
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   161
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   150
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   160
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   149
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   159
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   148
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   158
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   147
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   157
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   146
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   156
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   145
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   155
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   144
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   154
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   143
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   153
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   142
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   152
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   141
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   151
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   140
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   150
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   139
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   149
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   138
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   148
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   137
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   147
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   136
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   146
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   135
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   145
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   134
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   144
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   133
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   143
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   132
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   142
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   131
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   141
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   130
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   140
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   129
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   139
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   128
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   138
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   127
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   137
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   126
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   136
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   125
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   135
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   124
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   134
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   123
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   133
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   122
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   132
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   121
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   131
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   120
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   130
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   119
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   129
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   118
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   128
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   117
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   127
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   116
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   126
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   115
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   125
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   114
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   124
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   113
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   123
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   112
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   122
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   111
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   121
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   110
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   120
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   109
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   119
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   108
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   118
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   107
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   117
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   106
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   116
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   105
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   115
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   104
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   114
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   103
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   113
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   102
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   112
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   101
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   111
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   100
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   110
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   99
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   109
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   98
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   108
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   97
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   107
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   96
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   106
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   95
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   105
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   94
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   104
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   93
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   103
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   92
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   102
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   91
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   101
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   90
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   100
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   89
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   99
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   88
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   98
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   87
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   97
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   86
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   96
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   85
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   95
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   84
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   94
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   83
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   93
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   82
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   92
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   81
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   91
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   80
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   90
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   79
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   89
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   78
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   88
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   77
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   87
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   76
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   86
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   75
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   85
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   74
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   84
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   73
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   83
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   72
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   82
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   71
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   81
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   70
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   80
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   69
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   79
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   68
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   78
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   67
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   77
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   66
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   76
         Top             =   1920
         Width           =   495
         Begin MSComDlg.CommonDialog C 
            Left            =   240
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   65
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   75
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   64
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   74
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   63
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   73
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   62
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   72
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   61
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   71
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   60
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   70
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   59
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   69
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   58
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   68
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   57
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   67
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   56
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   66
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   55
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   65
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   54
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   64
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   53
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   63
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   52
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   62
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   51
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   61
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   50
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   60
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   49
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   59
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   48
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   58
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   47
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   57
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   46
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   56
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   45
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   55
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   54
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   53
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   52
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   51
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   50
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   49
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   48
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   47
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   46
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   45
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   44
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   43
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   42
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   41
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   40
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   39
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   36
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   35
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   28
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   6720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   25
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   6240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   5280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   22
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   21
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   17
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   12
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picstore 
      Height          =   495
      Left            =   7440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Pieces"
      Height          =   7335
      Left            =   7440
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.VScrollBar VScroll1 
         Height          =   2055
         Index           =   2
         Left            =   2520
         Max             =   1500
         Min             =   1
         TabIndex        =   276
         Top             =   5160
         Value           =   750
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2175
         Index           =   1
         Left            =   2520
         Max             =   500
         Min             =   1
         TabIndex        =   275
         Top             =   2760
         Value           =   200
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2175
         Index           =   0
         Left            =   2520
         Max             =   900
         Min             =   5
         TabIndex        =   272
         Top             =   360
         Value           =   50
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   720
         Picture         =   "3dmapeditor.frx":0442
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   271
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   5520
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   120
         Picture         =   "3dmapeditor.frx":1084
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   270
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   5520
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   1320
         Picture         =   "3dmapeditor.frx":1CC6
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   269
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   5520
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   1920
         Picture         =   "3dmapeditor.frx":2908
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   268
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   5520
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   720
         Picture         =   "3dmapeditor.frx":354A
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   267
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6120
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   120
         Picture         =   "3dmapeditor.frx":418C
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   266
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6120
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   1320
         Picture         =   "3dmapeditor.frx":4DCE
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   265
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6120
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   1920
         Picture         =   "3dmapeditor.frx":5A10
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   264
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6120
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   720
         Picture         =   "3dmapeditor.frx":6652
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   263
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   120
         Picture         =   "3dmapeditor.frx":7294
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   262
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   1320
         Picture         =   "3dmapeditor.frx":7ED6
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   261
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   1920
         Picture         =   "3dmapeditor.frx":8B18
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   260
         ToolTipText     =   "Beach/Grass Edge"
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   1920
         Picture         =   "3dmapeditor.frx":975A
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   259
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   1320
         Picture         =   "3dmapeditor.frx":A39C
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   258
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   120
         Picture         =   "3dmapeditor.frx":AFDE
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   257
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   720
         Picture         =   "3dmapeditor.frx":BC20
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   256
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   1920
         Picture         =   "3dmapeditor.frx":C862
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   255
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   1320
         Picture         =   "3dmapeditor.frx":D4A4
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   254
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   120
         Picture         =   "3dmapeditor.frx":E0E6
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   253
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   720
         Picture         =   "3dmapeditor.frx":ED28
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   252
         ToolTipText     =   "Water/Grass Edge"
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   1920
         Picture         =   "3dmapeditor.frx":F96A
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   251
         ToolTipText     =   "Water/Beach Edge"
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   1320
         Picture         =   "3dmapeditor.frx":105AC
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   250
         ToolTipText     =   "Water/Beach Edge"
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   120
         Picture         =   "3dmapeditor.frx":111EE
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   249
         ToolTipText     =   "Water/Beach Edge"
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   720
         Picture         =   "3dmapeditor.frx":11E30
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   248
         ToolTipText     =   "Water/Beach Edge"
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   1920
         Picture         =   "3dmapeditor.frx":12A72
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   247
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   1320
         Picture         =   "3dmapeditor.frx":136B4
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   246
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   120
         Picture         =   "3dmapeditor.frx":142F6
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   245
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   720
         Picture         =   "3dmapeditor.frx":14F38
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   244
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   1920
         Picture         =   "3dmapeditor.frx":15B7A
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   243
         ToolTipText     =   "Rock"
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   120
         Picture         =   "3dmapeditor.frx":167BC
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   242
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   720
         Picture         =   "3dmapeditor.frx":173FE
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   241
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   1320
         Picture         =   "3dmapeditor.frx":18040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   240
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   1920
         Picture         =   "3dmapeditor.frx":18C82
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   239
         ToolTipText     =   "Water/Beach Edge"
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   720
         Picture         =   "3dmapeditor.frx":198C4
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   238
         ToolTipText     =   "Path"
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   120
         Picture         =   "3dmapeditor.frx":1A506
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   237
         ToolTipText     =   "Beach"
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   1320
         Picture         =   "3dmapeditor.frx":1B148
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   236
         ToolTipText     =   "Tree"
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   1920
         Picture         =   "3dmapeditor.frx":1BD8A
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   8
         ToolTipText     =   "Water/Grass Edge"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   1320
         Picture         =   "3dmapeditor.frx":1C9CC
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   7
         ToolTipText     =   "Water/Grass Edge"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   120
         Picture         =   "3dmapeditor.frx":1D60E
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   6
         ToolTipText     =   "Water/Grass Edge"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   720
         Picture         =   "3dmapeditor.frx":1E250
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         ToolTipText     =   "Water/Grass Edge"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         Picture         =   "3dmapeditor.frx":1EE92
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         ToolTipText     =   "Bridge"
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1320
         Picture         =   "3dmapeditor.frx":1FAD4
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   3
         ToolTipText     =   "Bridge"
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   720
         Picture         =   "3dmapeditor.frx":20716
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   2
         ToolTipText     =   "Water"
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         FillColor       =   &H0000C0C0&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "3dmapeditor.frx":21358
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   1
         ToolTipText     =   "Grass"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7.5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   280
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   279
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Turning"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   278
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   277
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   274
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   273
         Top             =   1080
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   2400
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   2400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2400
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu compile 
         Caption         =   "&Compile 3D"
         Shortcut        =   ^C
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu reset 
         Caption         =   "&Reset"
         Shortcut        =   ^R
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim current As Integer


Private Sub compile_Click()
Module1.height = Label2(0).Caption
Module1.speed = Label2(1).Caption
Module1.turning = Label2(2).Caption


Dim start, num, counter, count2, square, coin As Integer
square = 1
coin = 0
count2 = 1
start = 1
z(0) = 0
X(0) = 0

For counter = 1 To 225
DoEvents
z(square) = coin
square = square + 15
count2 = count2 + 1
If count2 = 16 Then
start = start + 1
square = start
coin = coin + 200
count2 = 1
End If
Next counter

num = 0
count2 = 1

For counter = 1 To 225
DoEvents
X(counter) = num
count2 = count2 + 1
If count2 = 16 Then
count2 = 1
num = num + 200
End If
Next counter


For counter = 1 To 225
For count2 = 0 To 43
DoEvents
If Picture1(counter).Picture = Picture2(count2).Picture Then tex(counter) = count2
Next count2
Next counter


frmMain.Show
End Sub

Private Sub exit_Click()
If MsgBox("Are you sure want to exit?", vbYesNo, "Map Editor") = vbNo Then Exit Sub
End
End Sub

Private Sub Form_Load()

Dim counter As Integer
current = 0
For counter = 1 To 225
DoEvents
Picture1(counter).Picture = Picture2(0).Picture
Next counter
picstore.Picture = Picture2(0).Picture
End Sub



Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure want to exit?", vbYesNo, "Map Editor") = vbYes Then
End
Else
Cancel = 1
End If
End Sub

Private Sub open_Click()
Dim counter, count2 As Integer
Dim var
Form1.MousePointer = vbHourglass
    C.flags = cdlOFNFileMustExist
    C.Filter = "Data Files (*.dat)|*.dat"
    C.DialogTitle = "Open"
    C.ShowOpen
    If C.FileName = "" Then
        Form1.MousePointer = 0
        Exit Sub
    End If
Open C.FileName For Input As #1
For counter = 1 To 225
Line Input #1, var
Picture1(counter).Picture = Picture2(var).Picture
Next counter
Line Input #1, var
VScroll1(0).Value = var
Line Input #1, var
VScroll1(1).Value = var
Line Input #1, var
VScroll1(2).Value = var
Close #1
Form1.MousePointer = 0
End Sub

Private Sub Picture1_Click(Index As Integer)
Picture1(Index).Picture = picstore.Picture
End Sub

Private Sub Picture2_Click(Index As Integer)
picstore.Picture = Picture2(Index).Picture
Picture2(current).Appearance = 0
Picture2(Index).Appearance = 1
current = Index
End Sub

Private Sub reset_Click()
If MsgBox("Are you sure want to clear the map?", vbYesNo, "Map Editor") = vbNo Then Exit Sub
Dim counter As Integer
current = 0
For counter = 1 To 225
DoEvents
Picture1(counter).Picture = Picture2(0).Picture
Next counter
picstore.Picture = Picture2(0).Picture
End Sub

Private Sub save_Click()
Dim counter, count2, var
Form1.MousePointer = vbHourglass
    C.flags = cdlOFNOverwritePrompt
    C.Filter = "Data Files (*.dat)|*.dat"
    C.DialogTitle = "Save As"
    C.ShowSave
    If C.FileName = "" Then
        Form1.MousePointer = 0
        Exit Sub
    End If
Open C.FileName For Output As #1
For counter = 1 To 225
DoEvents
For count2 = 0 To 43
DoEvents
If Picture1(counter).Picture = Picture2(count2).Picture Then var = count2
Next count2
var = "0" & var
Print #1, var
Next counter
Print #1, VScroll1(0).Value
Print #1, VScroll1(1).Value
Print #1, VScroll1(2).Value
Close #1
Form1.MousePointer = 0
End Sub


Private Sub VScroll1_Change(Index As Integer)
Label2(0).Caption = VScroll1(0).Value
Label2(1).Caption = VScroll1(1).Value
Label2(2).Caption = VScroll1(2).Value / 100




End Sub
