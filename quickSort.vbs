'XX is the array to test QuickSort function
'Author: xiaomingneo
'----------------------------
Const N=20
Dim XX(20)
'Generate 20 random integer as the test array
For I=0 To N-1
	Randomize
	XX(I)=CInt(Rnd*100)
Next

ss="Before£º"&vbNewLine
For I = 0 To N
	ss = ss & XX(I) & "  "
Next
MsgBox ss
ss=ss & vbNewLine & "After QuickSort()£º" & vbNewLine


temp = QuickSort(XX,0,N)
For I = 0 To N
	ss = ss & XX(I) & "  "
Next
MsgBox ss

'------------------------------------
'Quick Sort Function
'take the first element of the array as the pivot
'composed by three steps:
' 1. Partition by pivot number.
' 2. Sort elements before pivot
' 3. Sort elements after pivot
'------------------------------------
Function QuickSort(data,low,high)
	Dim pivotpos
	If low<high Then
		pivotpos=Partition(data,low,high)
		temp = QuickSort(data,low,pivotpos-1)
		temp = QuickSort(data,pivotpos+1,high)
	End If
	QuickSort=0
End Function

'---------------------
' After Partition, the elements before pivot are smaller than pivot
' elements after pivot are all larger than pivot
'---------------------
Function Partition(data,byval i,byval j)
	pivot = data(i)
	Do
		Do While i<j And data(j)>=pivot
			j=j-1	
		Loop
		If i<j Then
			data(i)=data(j)
			i=i+1
		End If
		Do While i<j And data(i)<pivot
			i=i+1	
		Loop
		If i<j Then
			data(j)=data(i)
			j=j-1
		End If
	Loop While i<j
	data(i)=pivot
	Partition=i
End Function