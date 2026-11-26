'========================================================================================
' SMARTABSENT - APLIKASI PENCATATAN KEHADIRAN
'========================================================================================
'Nama	:Seyna Pasang
'Kelas	:XII-1
'Tanggal:Kamis 15 Oktober 2026
'========================================================================================

Option Explicit

Dim namaSiswa
Dim tanggal
Dim pilihan
Dim status
Dim catatan

Dim hadir
Dim izin
Dim sakit
Dim alpa
Dim total

'========================================================================================
'NILAI AWAL
'========================================================================================
 
 hadir = 0
 izin = 0
 sakit = 0
 alpa = 0
 
'Mengambil tanggal komputer secara otomatis
 tanggal = Date

'========================================================================================
'DATA SISWA
'========================================================================================

namaSiswa = InputBox(" SMARTABSENT " & vbCrLf & _
"APLIKASI PENCATATAN KEHADIRAN" & vbCrLf & vbCrLf & _
 "Masukkan nama siswa:", _
  "DATA SISWA")
  
  If namaSiswa = "" Then
  
MsgBox "Nama siswa belum diisi.", vbExclamation, "Peringatan"
WScript.Quit
End If

'========================================================================================
'MENU UTAMA
'========================================================================================
Do 
total = hadir + izin + sakit + alpa 

pilihan = InputBox(" SMATERABSENT " & vbCrLf & vbCrLf & _
"Nama		: " & namaSiswa & vbCrLf & _ 
"Tanggal		: " & tanggal & vbCrLf & _ 
"Total catatan	: " & total & vbCrLf & vbCrLf & _
"______________________________________________________________" & vbCrLf & _ 
"1. Catatan Kehadiran" & vbCrLf & _ 
"2. Lihat Rekap Kehadiran" & vbCrLf & _ 
"3. Keluar" & vbCrLf& _ 
"______________________________________________________________" & vbCrLf & _ 
"Masukkan pilihan:", "MENU UTAMA")
 
Select Case pilihan:

Case "1"
CatatanKehadiran

Case "2"
LihatRekap

Case "3"
MsgBox "Sampai jumpa!.", vbInformation, "Selesai"

Case Else
If pilihan <> "" Then
MsgBox "Pilihan tidak tersedia.", vbExclamation, "Peringatan"
End If
End Select
Loop While pilihan <> "3" And Pilihan <> ""


'========================================================================================
'SUB PROGRAM CATATAN KEHADRAN
'========================================================================================

Sub CatatanKehadiran()
status = InputBox(" CATAT KEHADIRAN " & vbCrLf & vbCrLf & _ 
"Nama	: " & namaSiswa & vbCrLf & _ 
"Tanggal: " & tanggal & vbCrLf & vbCrLf & _ 
"1. Hadir" & vbCrLf & _ 
"2. Izin" & vbCrLf & _ 
"3. Sakit" & vbCrLf & _ 
"4. Alpa" & vbCrLf & _ 
"Masukkan nomor status:", _
"STATUS KEHADIRAN")

Select Case status

Case "1"
hadir = hadir + 1
catatan = "Hadir"

Case "2"
izin = izin + 1
catatan = "Izin"

Case "3"
sakit = sakit + 1
catatan = "Sakit"

Case "4"
alpa = alpa + 1
catatan = "Alpa"

Case ""
Exit Sub

Case Else 
MsgBox "Status tidak valid!", vbExclamation, "Peringatan"
Exit Sub
End Select

MsgBox "Kehadiran berhadil dicatat!" & vbCrLf & vbCrLf & _ 
"Nama	: " & namaSiswa & vbCrLf & _ 
"Tanggal: " & tanggal & vbCrLf & _ 
"status	: " & catatan, vbInformation, "Berhasil"
End Sub

'========================================================================================
'SUB PROGRAM REKAP KEHADIRAN
'========================================================================================

Sub LihatRekap()
total = hadir + izin + sakit + alpa

MsgBox " REKAP KEHADIRAN " & vbCrLf & _ 
"Nama Siswa: " & namaSiswa & vbCrLf & _ 
"Tanggal	: " & tanggal & vbCrLf & _ 
"______________________________________________________________" & vbCrLf & _ 
"Hadir	: " & hadir &" kali" & vbCrLf & _ 
"Izin	: " & izin &" kali" & vbCrLf & _ 
"Sakit	: " & sakit &" kali" & vbCrLf & _ 
"Alpa	: " & alpa &" kali" & vbCrLf & _ 
"______________________________________________________________" & vbCrLf & _ 
"Total	: " & total &" kali", vbInformation, "Rekap"
End Sub

'========================================================================================
'Mematikan komputer, dengan komfirmasi dulu supaya tidak sengaja
'========================================================================================
Sub MatikanKomputer()
Dim jawaban, shell
jawaban = MsgBox("Yakin ingin mematikan komputer sekarang?", vbYesNo + vbQuestion, "Konfirmasi")
If jawaban = vbYes Then
Set shell = Createobject("WScript.Shell")
shell.Run "shutdown /s /t 5", 0, True
MsgBox "Komputer akan mati dalam 5 detik.", vbInformation, "Mematikan..."
Else 
MsgBox "Dibatalkan.", vbInformation, "Batal"
End If
End Sub 