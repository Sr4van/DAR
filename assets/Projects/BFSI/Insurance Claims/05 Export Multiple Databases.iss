Dim ndata$() AS string
Dim rdata$() AS string

Begin Dialog Dialog1 0,0,379,191,"Selection"
  ListBox 4,48,124,70, ndata$(), .ListBox1
  ListBox 165,48,124,70, rdata$(), .ListBox2
  Text 4,10,110,8, "Select databases to be exported.", .Text1
  Text 4,33,124,9, "IDEA databases found in the selected folder:", .Text2
  Text 165,34,124,10, "IDEA databases to be exported:", .Text3
  PushButton 135,48,20,15, ">", .PushButton1
  PushButton 135,65,20,15, "<", .PushButton2
  OKButton 90,122,40,14, "OK", .OKButton1
  CancelButton 152,122,40,14, "Cancel", .CancelButton1
  OptionGroup .option
  OptionButton 293,50,57,15, "Excel 2007/2010", .OptionButton1
  OptionButton 293,70,40,14, "PDF", .OptionButton2
  OptionButton 293,88,40,14, "Text File", .OptionButton3
  Text 302,35,27,9, "File Type:", .Text4
  Text 191,139,149,7, "Note: The Microsoft Excel format only allows up to 1,048,575 records per file.", .Text5
  PushButton 135,83,20,15, ">>", .PushButton3
  PushButton 135,100,20,15, "<<", .PushButton4
End Dialog









'***************************************************************************************************************************************************************
'***IDEAScript:    Export Multiple Databases.iss
'***Author:          Oscar Salas	
'***Date:             2011/04/14
'***Purpose:        Exports multiple IDEA databases and saves them in a selected folder. This macro provides the ability to simultaneously export several IDEA databases to 
'		     either Microsoft Excel, Adobe PDF, or Text format and save them to a folder of your choice.
'		Note: You must have IDEA open and have a project open. You will only be able to export databases from that project. 
'***Contact info: ideasupport@caseware.com
'***Disclaimer:   This script is provided as is without any warranties.
'***************************************************************************************************************************************************************
Sub Main
	Dim dlg As Dialog1
	Dim aux1()
	ans=MsgBox ("This macro lists all IDEA databases stored in a selected folder. From this list, select the databases to be exported and saved.",64,"Purpose of this Macro")
	outdir=BrowseFolder("Select the folder where the databases are located:")
	If outdir="" Then
		Exit Sub
	Else
        		outdir=outdir & "\"
	End If           
	edata=Dir(outdir & "*.IMD")
	If edata="" Then
      		ans=MsgBox ("The directory you selected does not contain any IDEA databases.",16,"Error")
      		Exit Sub
	Else
           		i=0
		Do Until edata=""
			edata=Dir()
			If edata<>"" Then
				i=i+1
			End If   
		Loop       
   	End If          
	ReDim ndata$(i)
	ReDim rdata$(i)
   	ReDim aux1(i)
	'Gets the file names of all IDEA databases and stores them in an array.
   	string1=Dir(outdir & "*.IMD")
   	lstring=Len(string1)
   	ndata$(0)=Mid(string1,1,lstring-4)
   	For j=1 To i
          		string1=Dir()
          		lstring=Len(string1)
          		ndata$(j)=Mid(string1,1,lstring-4)
   	Next i       
   	dlg.PushButton4=False
	b:   
		res=Dialog(dlg)
		'Displays the selection dialog box and moves objects between ListBox1 and ListBox2
		If res=1 Then               
      			If dlg.ListBox1=-1 Then
         				GoTo b
      			Else
              			aux=ndata$(dlg.ListBox1)
              			ndata$(dlg.ListBox1)=""
              			k=0
              			For j=0 To i
                     				If ndata$(j)<>"" Then
                        				aux1(k)=ndata$(j)
                        				ndata$(j)=""
                        				k=k+1
                     				End If
              			Next j
              			k=0
              			For j=0 To i
                     				If aux1(j)<>"" Then
                        				ndata$(k)=aux1(j)
                        				aux1(j)=""
                        				k=k+1
                     				End If
              			Next j
              			For j=0 To i
                     				If rdata$(j)="" Then
                        				rdata$(j)=aux
                        				Exit For
                     				End If   
              			Next j          
              			GoTo b
			End If        
   		ElseIf res=2 Then
			If dlg.ListBox2=-1 Then
                 			GoTo b
              		Else    
                      			aux=rdata$(dlg.ListBox2)
                      			rdata$(dlg.ListBox2)=""
                      			k=0
                      			For j=0 To i
                             			If rdata$(j)<>"" Then
                                				aux1(k)=rdata$(j)
                                				rdata$(j)=""
                                				k=k+1
                             			End If
                      			Next j
                      			k=0
                      			For j=0 To i
                             			If aux1(j)<>"" Then
                                				rdata$(k)=aux1(j)
                                				aux1(j)=""
                                				k=k+1
                             			End If
                      			Next j
                      			For j=0 To i
                             			If ndata$(j)="" Then
                                				ndata$(j)=aux
                                				Exit For
                             			End If
                      			Next j
                      			GoTo b
              		End If
		ElseIf res=3 Then
              		For j=0 To i
                    			If rdata$(j)="" Then
                        			For k=0 To i
                               				If ndata$(k)<>"" Then
                                  					rdata$(j)=ndata$(k)
                                  					ndata$(k)=""
                                  					Exit For
                               				End If   
                        			Next k
                     			End If
              		Next j          
              		GoTo b              
		ElseIf res=4 Then
              		For j=0 To i
                     			If ndata$(j)="" Then
                        			For k=0 To i
                               				If rdata$(k)<>"" Then
                                  					ndata$(j)=rdata$(k)
                                  					rdata$(k)=""
                                  					Exit For
                               				End If
                        			Next k
                     			End If
              		Next j
			GoTo b                    
		ElseIf res=0 Then 
			Exit Sub
		ElseIf res=-1 Then 
              		If rdata$(0)="" Then
                 			ans=MsgBox ("Please select a database to export.",16,"Invalid selection")
                 			GoTo b
              		End If
              		k=0
              		j=0
              		Do Until j=i+1
                     			If rdata$(j)<>"" Then
                        			k=k+1       
                     			End If   
                     			j=j+1
              		Loop
			outdir1=BrowseFolder("Select the folder to save the files to:")
              		If outdir1="" Then
                 			Exit Sub
              		End If   
              		outdir1=outdir1 & "\"
              		choice=dlg.option
              		Set pcompl=CreateObject("CommonIdeaControls.StandaloneProgressCtl")
                     		pcompl.Start "Processing IDEA databases......."
                     		stp=100/k
              		Select Case choice
				Case 0
					ext="XLSX"
					ext1=".XLSX"
					For j=0 To k-1
						Set db=Client.OpenDatabase(outdir & rdata$(j) & ".IMD")
				            	pcompl.Progress stp * j
				       		If db.Count=0 Then
				       		Else
				            		If db.Count>1048575 Then
				                  			records=1048575
				               		Else
				                      			records=db.Count
				               		End If           
				               		Set task=db.ExportDatabase
				                      		task.IncludeAllFields
				                      		task.PerformTask outdir1 & rdata$(j) & ext1,"Database",ext,1,records,""
				               		Set task=Nothing
				               		Set db=Nothing
				               		Client.CloseAll
						End If         
					Next j
					pcompl.Progress 100
				Case 1
					ext1=".pdf"
					For j=0 To k-1
				       		Set db=Client.OpenDatabase(outdir & rdata$(j) & ".IMD")
				              	pcompl.Progress stp * j
				       		If db.Count=0 Then
				       
				       		Else       
				               		db.FileNameForPublishing=outdir1 & rdata$(j) & ext1
				               		db.PublishToPDF
				       			Set db=Nothing
				       			Client.CloseAll
				       			End If
					Next j
					pcompl.Progress 100
				Case 2
					ext="ASC"
					ext1=".ASC"
					For j=0 To k-1
				       		Set db=Client.OpenDataBase(outdir & rdata$(j) & ".IMD")
				       		pcompl.Progress stp * j
				       		Set task=db.ExportDatabase
				              	task.IncludeAllFields
				              	eqn=""
				             	 task.Separators ",", "."
				              	task.PerformTask outdir1 & rdata$(j) & ext1,rdata$(j),ext,1,db.Count,eqn
				       		Set task=Nothing
				       		Set db=Nothing
				       		Client.CloseAll
					Next j
					pcompl.Progress 100
              			End Select
		End If
		ans=MsgBox ("Process Completed",64,"Information")                                               
End Sub
'**************************************************************************************************************************************
'***Function Name: BrowseFolder
'***Purpose:	         This function is for browsing through a folder 
'**************************************************************************************************************************************
Function BrowseFolder(title As String)
	Dim oFolder
	Dim oFolderItem
	Dim oPath
	Dim oShell
	Dim strPath
	Dim strCurDir
	Set oShell = CreateObject( "Shell.Application" )
	Set strCurDir = Client.WorkingDirectory
	Set oFolder = oShell.BrowseForFolder(0, title, 1,strCurDir)
	If oFolder Is Nothing Then
		BrowseFolder = ""
   		Exit Sub
	End If
	Set oFolderItem = oFolder.Self
	BrowseFolder=oFolderItem.Path
End Function

