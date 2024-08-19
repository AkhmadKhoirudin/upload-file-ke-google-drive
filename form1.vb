Imports Google.Apis.Auth.OAuth2   'intasll manual di NuGet
Imports Google.Apis.Drive.v3      'intasll manual di NuGet
Imports Google.Apis.Services
Imports Google.Apis.Drive.v3.Data
Imports System.IO
Imports System.Threading

Public Class Form1
    Private filePaths As New List(Of String)()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "All Files|*.*"
            openFileDialog.Title = "Pilih File"
            openFileDialog.Multiselect = True ' Mengaktifkan pemilihan multi-file

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                filePaths.Clear()
                ListBox1.Items.Clear()
                filePaths.AddRange(openFileDialog.FileNames)
                For Each filePath In filePaths
                    ListBox1.Items.Add(Path.GetFileName(filePath))
                Next
            End If
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If filePaths.Count = 0 Then
            MessageBox.Show("Silakan pilih file terlebih dahulu.")
            Return
        End If

        ' Path ke file JSON kredensial
        Dim credentialsPath As String = "client_secret_207258522611-5lif41l50g00ggnmag8p9h5lqhj02e6k.apps.googleusercontent.com.json"

        ' Load OAuth 2.0 credentials
        Dim credential As UserCredential
        Using stream = New FileStream(credentialsPath, FileMode.Open, FileAccess.Read)
            Dim secrets = GoogleClientSecrets.FromStream(stream).Secrets
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                secrets,
                {DriveService.Scope.Drive},
                "user",
                CancellationToken.None).Result
        End Using

        ' Membuat Drive API service
        Dim service = New DriveService(New BaseClientService.Initializer() With {
            .HttpClientInitializer = credential,
            .ApplicationName = "Your Application Name"
        })

        ProgressBar1.Maximum = filePaths.Count
        ProgressBar1.Value = 0
        ProgressBar1.Step = 1

        Dim allUploadsSuccessful As Boolean = True ' Flag untuk melacak status upload

        For Each filePath In filePaths
            ' Metadata file yang akan diunggah
            Dim fileMetadata = New Google.Apis.Drive.v3.Data.File() With {
                .Name = Path.GetFileName(filePath),
                .Parents = New List(Of String) From {"11BYoSgJJLZBizzuA6W3gpqwym_d3n8-U"} ' ID folder tujuan
            }
            Dim fileStream = New FileStream(filePath, FileMode.Open)

            Dim request = service.Files.Create(fileMetadata, fileStream, GetMimeType(filePath))
            request.Fields = "id"
            Dim file = request.Upload()

            If file.Status = Google.Apis.Upload.UploadStatus.Completed Then
                ProgressBar1.PerformStep()
            Else
                MessageBox.Show("Error mengunggah file: " & Path.GetFileName(filePath))
                allUploadsSuccessful = False ' Set flag ke false jika ada upload yang gagal
            End If

            fileStream.Close()
        Next

        ' Kosongkan ListBox dan tampilkan pesan sukses jika semua file berhasil diunggah
        If allUploadsSuccessful Then
            ListBox1.Items.Clear()
            MessageBox.Show("Semua file berhasil diunggah.")
        Else
            MessageBox.Show("Beberapa file gagal diunggah.")
        End If
    End Sub

    Private Function GetMimeType(filePath As String) As String
        Dim extension As String = Path.GetExtension(filePath).ToLowerInvariant()
        Select Case extension
            Case ".jpg", ".jpeg"
                Return "image/jpeg"
            Case ".png"
                Return "image/png"
            Case ".pdf"
                Return "application/pdf"
            Case ".doc", ".docx"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case Else
                Return "application/octet-stream"
        End Select
    End Function
End Class
