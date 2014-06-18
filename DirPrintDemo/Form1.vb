Public Class Form1

    Const IMAGE_FOLDER_CLOSED = 0
    Const IMAGE_FOLDER_OPEN = 1
    Const IMAGE_FILE = 2
    Const IMAGE_DRIVE = 3

    Dim printNode As TreeNode = Nothing
    Dim printLevel As Integer = 0

    Dim printDirStack As Stack(Of String)
    Dim printFiles() As String
    Dim printFilesIndex As Integer

    Enum enumNodeType
        Drive
        Folder
        File
        Empty
    End Enum


    Private Class classNodeValues
        Friend nodeType As enumNodeType
        Friend fullPath As String

        Friend fileSize As Long
        Friend fileDate As Date

        Public Sub New(ByVal node As enumNodeType, _
                       ByVal path As String)
            nodeType = node
            fullPath = path
            fileSize = 0
            fileDate = Date.MinValue
        End Sub

        Public Sub New(ByVal node As enumNodeType, _
                       ByVal path As String, _
                       ByVal fSize As Long, _
                       ByVal fDate As Date)
            nodeType = node
            fullPath = path
            fileSize = fSize
            fileDate = fDate
        End Sub

    End Class

    Private Sub textFolder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textFolder.TextChanged
        PopulateFolderTree(textFolder.Text)
    End Sub

    Private Sub PopulateFolderTree(ByVal rootPath As String)
        TreeView1.Nodes.Clear()

        If Not String.IsNullOrEmpty(rootPath) Then
            If IO.Directory.Exists(rootPath) Then
                Dim entries() As String
                Dim node As TreeNode
                Dim folderName As String
                Dim rootPathLen As Integer
                Dim info As IO.DirectoryInfo
                Dim attributeMask As IO.FileAttributes
                Dim tmpNode As TreeNode

                attributeMask = IO.FileAttributes.Hidden Or IO.FileAttributes.System

                rootPathLen = rootPath.Length
                If Not rootPath.EndsWith("\") Then
                    rootPathLen = rootPathLen + 1
                End If

                Try
                    entries = IO.Directory.GetDirectories(rootPath)
                    Array.Sort(entries)
                    For Each entry As String In entries
                        info = New IO.DirectoryInfo(entry)
                        If (info.Attributes And attributeMask) = 0 Then
                            ' extract just the folder name from the directory path
                            folderName = entry.Substring(rootPathLen, entry.Length - rootPathLen)

                            node = TreeView1.Nodes.Add(folderName)
                            node.ImageIndex = IMAGE_FOLDER_CLOSED
                            node.SelectedImageIndex = IMAGE_FOLDER_CLOSED
                            node.Tag = New classNodeValues(enumNodeType.Folder, entry)

                            tmpNode = node.Nodes.Add("[EMPTY]")
                            tmpNode.Tag = New classNodeValues(enumNodeType.Empty, String.Empty)
                        End If
                    Next

                    entries = IO.Directory.GetFiles(rootPath)
                    Array.Sort(entries)
                    For Each entry As String In entries
                        node = TreeView1.Nodes.Add(IO.Path.GetFileName(entry))
                        node.ImageIndex = IMAGE_FILE
                        node.SelectedImageIndex = IMAGE_FILE

                        Dim f As IO.FileInfo
                        f = New IO.FileInfo(entry)
                        node.Tag = New classNodeValues(enumNodeType.File, entry, f.Length, f.LastWriteTime)
                    Next
                Catch ex As Exception

                End Try
            End If
        Else
            Dim entries() As System.IO.DriveInfo
            Dim node As TreeNode
            Dim tmpNode As TreeNode
            Dim label As String

            entries = System.IO.DriveInfo.GetDrives()
            For Each entry As System.IO.DriveInfo In entries
                Try
                    If String.IsNullOrEmpty(entry.VolumeLabel) Then
                        label = entry.Name
                    Else
                        label = entry.VolumeLabel & " (" & entry.Name & ")"
                    End If
                Catch ex As Exception
                    ' probably "drive not ready" exception
                    label = entry.Name
                End Try

                node = TreeView1.Nodes.Add(label)
                node.ImageIndex = IMAGE_DRIVE
                node.SelectedImageIndex = IMAGE_DRIVE
                node.Tag = New classNodeValues(enumNodeType.Drive, entry.Name)

                tmpNode = node.Nodes.Add("[EMPTY]")
                tmpNode.Tag = New classNodeValues(enumNodeType.Empty, String.Empty)
            Next
        End If
    End Sub

    Private Sub buttonFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonFolder.Click
        FolderBrowserDialog1.SelectedPath = textFolder.Text
        FolderBrowserDialog1.ShowNewFolderButton = False
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            textFolder.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub TreeView1_AfterCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCollapse
        Dim nodeValue As classNodeValues

        nodeValue = e.Node.Tag
        If nodeValue.nodeType = enumNodeType.Folder Then
            e.Node.ImageIndex = IMAGE_FOLDER_CLOSED
            e.Node.SelectedImageIndex = IMAGE_FOLDER_CLOSED
        End If
    End Sub

    Private Sub TreeView1_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        Dim nodeValue As classNodeValues
        Dim tmpValue As classNodeValues

        nodeValue = e.Node.Tag

        If e.Node.Nodes.Count = 1 Then
            tmpValue = e.Node.Nodes(0).Tag
            If tmpValue.nodeType = enumNodeType.Empty Then
                ' need to populate this branch before expansion
                e.Node.Nodes.Clear()

                Dim entries() As String
                Dim node As TreeNode
                Dim folderName As String
                Dim pathLen As Integer
                Dim info As IO.DirectoryInfo
                Dim attributeMask As IO.FileAttributes
                Dim tmpNode As TreeNode

                attributeMask = IO.FileAttributes.Hidden Or IO.FileAttributes.System

                pathLen = nodeValue.fullPath.Length
                If Not nodeValue.fullPath.EndsWith("\") Then
                    pathLen = pathLen + 1
                End If


                Try
                    entries = IO.Directory.GetDirectories(nodeValue.fullPath)
                    Array.Sort(entries)
                    For Each entry As String In entries
                        ' extract just the folder name from the directory path
                        folderName = entry.Substring(pathLen, entry.Length - pathLen)

                        info = New IO.DirectoryInfo(entry)
                        If (info.Attributes And attributeMask) = 0 Then

                            node = e.Node.Nodes.Add(folderName)
                            node.ImageIndex = IMAGE_FOLDER_CLOSED
                            node.SelectedImageIndex = IMAGE_FOLDER_CLOSED
                            node.Tag = New classNodeValues(enumNodeType.Folder, entry)

                            tmpNode = node.Nodes.Add("[EMPTY]")
                            tmpNode.Tag = New classNodeValues(enumNodeType.Empty, String.Empty)
                        End If
                    Next

                    entries = IO.Directory.GetFiles(nodeValue.fullPath)
                    Array.Sort(entries)
                    For Each entry As String In entries
                        node = e.Node.Nodes.Add(IO.Path.GetFileName(entry))
                        node.ImageIndex = IMAGE_FILE
                        node.SelectedImageIndex = IMAGE_FILE

                        Dim f As IO.FileInfo
                        f = New IO.FileInfo(entry)
                        node.Tag = New classNodeValues(enumNodeType.File, entry, f.Length, f.LastWriteTime)
                    Next
                Catch ex As Exception

                End Try
            End If
        End If

        If nodeValue.nodeType = enumNodeType.Folder Then
            If e.Node.Nodes.Count > 0 Then
                e.Node.ImageIndex = IMAGE_FOLDER_OPEN
                e.Node.SelectedImageIndex = IMAGE_FOLDER_OPEN
            Else
                e.Node.ImageIndex = IMAGE_FOLDER_CLOSED
                e.Node.SelectedImageIndex = IMAGE_FOLDER_CLOSED
            End If
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PopulateFolderTree(String.Empty)
    End Sub

    Private Sub TreeView1_DrawNode(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawTreeNodeEventArgs) Handles TreeView1.DrawNode
        Dim brushFore As Brush
        Dim brushBack As Brush
        Dim nodeValue As classNodeValues

        nodeValue = e.Node.Tag
        If nodeValue.nodeType = enumNodeType.File Then
            If e.Bounds.Left >= 0 Then
                If (e.State And TreeNodeStates.Selected) <> 0 Then
                    brushFore = New SolidBrush(SystemColors.HighlightText)
                    brushBack = New SolidBrush(SystemColors.ActiveCaption)
                Else  ' node not selected, use forecolor
                    If e.Node.ForeColor.Equals(Color.Empty) Then
                        ' no forecolor is set for the node, use treeview's forecolor
                        brushFore = New SolidBrush(TreeView1.ForeColor)
                    Else
                        brushFore = New SolidBrush(e.Node.ForeColor)
                    End If

                    If e.Node.BackColor.Equals(Color.Empty) Then
                        ' no backcolor is set for the node, use treeview's forecolor
                        brushBack = New SolidBrush(TreeView1.BackColor)
                    Else
                        brushBack = New SolidBrush(e.Node.BackColor)
                    End If
                End If

                ' clear the background
                e.Graphics.FillRectangle(brushBack, e.Bounds.Left, e.Bounds.Top, TreeView1.ClientRectangle.Width, e.Bounds.Height)

                DrawFile(e.Node.Text, nodeValue.fileSize, nodeValue.fileDate, _
                         e.Bounds.Left + 1, TreeView1.ClientRectangle.Width - 5, _
                         e.Bounds.Top + 1, e.Graphics, TreeView1.Font, brushFore)

                brushFore.Dispose()
                brushBack.Dispose()
            End If
            e.DrawDefault = False
        Else
            ' nothing special about this node, let the system draw it
            e.DrawDefault = True
        End If
    End Sub

    Private Sub TreeView1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TreeView1.SizeChanged
        TreeView1.Refresh()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        SetupPrinting()

        PrintDialog1.AllowCurrentPage = False
        PrintDialog1.AllowPrintToFile = True
        PrintDialog1.AllowSelection = False
        PrintDialog1.AllowSomePages = False

        ' no, setting PrintDialog1.Document to PrintDocument1 doesn't automatically cause the print dialog control
        ' to invoke the printing.  Need to invoke the Print method manually.
        'PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.Print()
        End If

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        SetupPrinting()

        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    ' sample code for transversing and printing the nodes of a tree view control
    '
    'Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
    '    Dim linesPerPage As Single = 0
    '    Dim yPos As Single = 0
    '    Dim count As Integer = 0
    '    Dim line As String = Nothing

    '    ' Calculate the number of lines per page.
    '    linesPerPage = e.MarginBounds.Height / TreeView1.Font.GetHeight(e.Graphics)
    '    yPos = e.MarginBounds.Top

    '    ' Print each node of the tree
    '    While count < linesPerPage
    '        If printNode Is Nothing Then
    '            printNode = TreeView1.Nodes(0)
    '        Else
    '            If printNode.Nodes.Count > 0 Then
    '                printNode = printNode.Nodes(0)
    '                printLevel = printLevel + 1
    '            Else
    '                If printNode.NextNode IsNot Nothing Then
    '                    printNode = printNode.NextNode
    '                Else
    '                    printNode = printNode.Parent
    '                    printLevel = printLevel - 1

    '                    Do While printNode.NextNode Is Nothing
    '                        printNode = printNode.Parent
    '                        printLevel = printLevel - 1

    '                        If printNode Is Nothing Then
    '                            Exit While
    '                        End If
    '                    Loop

    '                    If printNode IsNot Nothing Then
    '                        printNode = printNode.NextNode
    '                    End If
    '                End If
    '            End If
    '        End If

    '        If printNode IsNot Nothing Then
    '            Dim value As classNodeValues
    '            value = printNode.Tag
    '            If value.nodeType <> enumNodeType.Empty Then
    '                e.Graphics.DrawString(printNode.Text, TreeView1.Font, SystemBrushes.ControlText, e.MarginBounds.X + (printLevel * 20), yPos)
    '                yPos = yPos + TreeView1.Font.GetHeight(g)

    '                count += 1
    '            End If
    '        Else
    '            count = linesPerPage + 1
    '        End If
    '    End While

    '    ' If more lines exist, print another page.
    '    If (printNode IsNot Nothing) Then
    '        e.HasMorePages = True
    '    Else
    '        e.HasMorePages = False

    '        printNode = Nothing
    '        printLevel = 0
    '    End If
    'End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = e.MarginBounds.Height / TreeView1.Font.GetHeight(e.Graphics)
        yPos = e.MarginBounds.Top

        ' Print each line of the file.
        While count < linesPerPage
            If (printFiles IsNot Nothing) AndAlso (printFilesIndex < printFiles.Length) Then
                ' in the process of printing the files for the current folder...
                ' get the file information for this file
                Dim f As IO.FileInfo
                f = New IO.FileInfo(printFiles(printFilesIndex))

                DrawFile(IO.Path.GetFileName(printFiles(printFilesIndex)), f.Length, f.LastWriteTime, _
                         e.MarginBounds.Left + (printLevel * 20), e.MarginBounds.Right, yPos, _
                         e.Graphics, TreeView1.Font, SystemBrushes.ControlText)

                ' increment line (vertical) position and line count
                yPos = yPos + TreeView1.Font.GetHeight(e.Graphics)
                count = count + 1

                printFilesIndex = printFilesIndex + 1
                If printFilesIndex = printFiles.Length Then
                    printFiles = Nothing
                    printLevel = printLevel - 1

                    ' add a "blank line"
                    yPos = yPos + TreeView1.Font.GetHeight(e.Graphics)
                    count = count + 1
                End If
            Else  ' no files to print, move on to a new directory
                Dim dirPath As String
                Dim entries() As String
                Dim boldFont As Font

                If printDirStack.Count > 0 Then
                    dirPath = printDirStack.Pop

                    If dirPath.Equals(":") Then
                        ' finished printing any subfolders and files for the folder, decrease the indentation level
                        printLevel = printLevel - 1
                    ElseIf dirPath.StartsWith(":") Then
                        ' moving back up the directory tree, retrieve the file listing for this folder
                        dirPath = dirPath.Substring(1, dirPath.Length - 1)

                        printFiles = IO.Directory.GetFiles(dirPath)
                        If printFiles.Length > 0 Then
                            Array.Sort(printFiles)
                            printFilesIndex = 0
                            printLevel = printLevel + 1
                        Else
                            printFiles = Nothing
                        End If
                        printDirStack.Push(":")
                    Else
                        ' push an entry onto the stack to that after processing any subfolders, the files of this folder will be processed
                        printDirStack.Push(":" & dirPath)
                        printLevel = printLevel + 1

                        Dim folderName As String
                        If (dirPath.Length <= 3) Or (dirPath.Equals(textFolder.Text)) Then
                            ' it's either the root directory, or a drive letter.  in either case, print the entire string
                            folderName = dirPath
                        Else  ' just print out the folder name, not the entire path
                            Dim i As Integer
                            i = dirPath.LastIndexOf("\") + 1
                            folderName = dirPath.Substring(i, dirPath.Length - i)
                        End If

                        boldFont = New Font(TreeView1.Font, FontStyle.Bold)
                        e.Graphics.DrawString(folderName, boldFont, SystemBrushes.ControlText, e.MarginBounds.Left + (printLevel * 20), yPos)
                        boldFont.Dispose()

                        ' increment line (vertical) position and line count
                        yPos = yPos + TreeView1.Font.GetHeight(e.Graphics)
                        count = count + 1

                        entries = IO.Directory.GetDirectories(dirPath)
                        If entries.Length > 0 Then
                            Array.Sort(entries)
                            Array.Reverse(entries)
                            For Each entry As String In entries
                                printDirStack.Push(entry)
                            Next
                        End If
                    End If
                Else  ' finished printing all the folders
                    count = linesPerPage + 1
                End If
            End If
        End While

        ' If more lines exist, print another page.
        If (printFiles IsNot Nothing) Or (printDirStack.Count > 0) Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub SetupPrinting()
        printDirStack = New Stack(Of String)
        If Not String.IsNullOrEmpty(textFolder.Text) Then
            printDirStack.Push(textFolder.Text)
        Else  ' no directory specified, do all the drives
            Dim entries() As System.IO.DriveInfo

            entries = System.IO.DriveInfo.GetDrives()
            'Array.Sort(entries)
            Array.Reverse(entries)  ' so they are pushed on the stack in reverse sequence
            For Each entry As System.IO.DriveInfo In entries
                printDirStack.Push(entry.Name)
            Next
        End If
        printFiles = Nothing
        printFilesIndex = 0
        printLevel = 0
    End Sub

    Private Sub DrawFile(ByVal fileName As String, _
                         ByVal fileLength As Long, _
                         ByVal fileDate As Date, _
                         ByVal marginLeft As Integer, _
                         ByVal marginRight As Integer, _
                         ByVal yPos As Integer, _
                         ByRef g As Graphics, _
                         ByRef f As Font, _
                         ByRef b As Brush)
        Dim prtStr As String
        Dim prtSize As SizeF
        Dim colLeft As Single
        Dim width As Integer

        width = marginRight - marginLeft

        ' working from right to left...
        colLeft = marginRight

        ' format the file datestamp
        prtStr = fileDate.ToString("    MM/dd/yyyy  hh:mm tt")
        ' draw the file time
        prtSize = g.MeasureString(prtStr, f, width)
        colLeft = colLeft - prtSize.Width
        g.DrawString(prtStr, f, b, colLeft, yPos)

        ' determine the format to display the filesize
        Dim i As Long
        If fileLength = 0 Then
            prtStr = "0 KB"
        Else
            prtStr = " KB"
            i = fileLength / 1024
            If i = 0 Then
                i = 1
            ElseIf i >= 1024 Then
                prtStr = " MB"
                i = i / 1024
                If i >= 1024 Then
                    prtStr = " GB"
                    i = i / 1024
                End If
            End If
            prtStr = i.ToString & prtStr
        End If

        ' draw the file size
        prtSize = g.MeasureString(prtStr, f, width)
        colLeft = colLeft - prtSize.Width
        g.DrawString(prtStr, f, b, colLeft, yPos)

        ' draw the file name
        g.DrawString(fileName, f, b, marginLeft, yPos)
    End Sub

End Class
