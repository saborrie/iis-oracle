' Reports the script engine answering the probes, so every measurement in a
' run is attributable to an exact engine build.
WScript.Echo ScriptEngine & " " & ScriptEngineMajorVersion & "." & _
	ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion
