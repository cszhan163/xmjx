<SCRIPT language="vbscript" runat="server">
'--------------------------------------------------------------------------------------------------------------------------------
'	检测系统服务器端使用的组件是否可以使用
'--------------------------------------------------------------------------------------------------------------------------------
class DetectCom
	private aProgId(), aNumber(), aSource(), aDescription()
	public ComCount, ErrCount

	private sub class_initialize()
	end sub

	private sub class_terminate()
	end sub

	public function DetectAll()
		dim i
		ComCount = 4					'进行测试的组件个数
		ErrCount = 0					'记录发生错误的组件个数(0,1,2,...)
		
		redim aProgId(ComCount - 1)			'Variant Array
		redim aNumber(ComCount - 1)
		redim aSource(ComCount - 1)
		redim aDescription(ComCount - 1)

		aProgId(0) = "Scripting.FileSystemObject"
		aProgId(1) = "ADODB.Connection"
		aProgId(2) = "ADODB.Recordset"
		aProgId(3) = "ADODB.Command"

		for i = 0 to ComCount - 1
			DetectCom aProgId(i), i
		next
	end function
	
	public function Detect(ProgId)
		redim aProgId(0)			'Variant Array
		redim aNumber(0)
		redim aSource(0)
		redim aDescription(0)
		ComCount = 1					'进行测试的组件个数
		ErrCount = 0					'记录发生错误的组件个数(0,1,2,...)
		aProgId(0) = ProgId
		
		DetectCom aProgId(0), 0
	end function
	
	public property get ProgId()
		if ComCount = 1 then
			ProgId = aProgId(0)
		else
			ProgId = aProgId
		end if
	end property
	
	public property get Number()
		if ComCount = 1 then
			Number = aNumber(0)
		else
			Number = aNumber
		end if
	end property

	public property get Source()
		if ComCount = 1 then
			Source = aSource(0)
		else
			Source = aSource
		end if
	end property

	public property get Description()
		if ComCount = 1 then
			Description = aDescription(0)
		else
			Description = aDescription
		end if
	end property

	private function DetectCom(ProgId, i)
		dim Obj
		on error resume next
		set Obj = Server.CreateObject(ProgId)
		
		if Err.number <> 0 then
			aNumber(i) = Err.Number
			aSource(i) = Err.Source
			aDescription(i) = Err.Description
			ErrCount = ErrCount + 1
		end if

		set Obj = nothing
		on error goto 0
	end function
end class

'--------------------------------------------------------------------------------------------------------------------------------
'	检测系统的数据库连接文件 DB.udl
'--------------------------------------------------------------------------------------------------------------------------------
class DetectUdl
	public FileName, ErrCount, Description
	
	private sub class_initialize()
		ErrCount = 0				'记录发生错误的个数(0,1)
		FileName = ""				'要侦测的文件名
	end sub

	private sub class_terminate()
	end sub
	
	public function Detect()
		dim Obj
		
		set Obj = new DetectCom
		
		Obj.Detect "Scripting.FileSystemObject"
		if Obj.ErrCount <> 0 then
			ErrCount = 1
			Description = "无法进行检测！ "& Obj.Description
		else
			Obj.Detect "ADODB.Connection"
			if Obj.ErrCount <> 0 then
				ErrCount = 1
				Description = "无法进行检测！ "& Obj.Description 
			end if
		end if
		
		set Obj = nothing
		
		if ErrCount = 0 then
			set Obj = Server.CreateObject("Scripting.FileSystemObject")
			
			if not Obj.FileExists(FileName) then
				ErrCount = 1
				on error resume next
				Obj.CreateTextFile FileName
				if Err.number <> 0 then
					Description = "数据库连接文件创建失败！ "& Err.Description 
				else
					Description = "请设置数据库连接！"
				end if
				on error goto 0
			else
				set Obj = nothing
				
				set Obj = Server.CreateObject("ADODB.Connection")
				
				on error resume next
				Obj.Open "File Name="& FileName
				if Err.number <> 0 then
					ErrCount = 1
					Description = Err.Description 
				end if
				on error goto 0
			end if
			set Obj = nothing
		end if
	end function
end class
</SCRIPT>