namespace eval activexscripthost {
    proc _load_dll {directory} {
        # escape DLL hell - change to directory containing
        # activexscripthost.dll temporarily because a type library
        # has to be loaded from there
        set wd [pwd]
        cd $directory
        # execute library loading in global namespace because the library
        # init function creates a subnamespace
        namespace eval :: {load activexscripthost[info sharedlib]}
        cd $wd
        # hide this command when the library was loaded
        rename _load_dll ""
    }
}
