Version 2:
	~ bug fix: 		- continuously loading defaults (or last) values
				- could start empty files creation, even with no "use word/excel" selection
				- while empty files running "use word/excel" could be changed
				- if word/excel operations finish before "empty files" creation, button remains at "STOP" state
					restoreAfterRun() do not run before empty files creation is also completed 

	~ add:			- if (autoDelete) deletes also "empty files" folder (if user doesn't choose to open 
					the folder containing the files)

	~ known issues:		- at manual STOP (or Empty creation stop), excel instance remains (sometimes)
				
				