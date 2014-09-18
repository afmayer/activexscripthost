package require activexscripthost

proc TestAll {} {
    foreach testproc [info commands activexscripthost::test::*] {
        if [catch {$testproc} result] {
            error "[namespace tail $testproc]: $result"
        }
    }
}

# ActiveX Script Host test procedures
namespace eval activexscripthost::test {

proc MultiDimVariantArrayToFlatList {} {
    set vb [::activexscripthost::openengine VBScript]
    set expectedValues \
        [list 0,0 1,0 "" 0,1 1,1 "" 0,2 1,2 "" "" "" "" "" "" ""]
    $vb parse {
        Dim array(2, 4)
        array(0, 0) = "0,0"
        array(0, 1) = "0,1"
        array(0, 2) = "0,2"
        array(1, 0) = "1,0"
        array(1, 1) = "1,1"
        array(1, 2) = "1,2"
    }
    $vb setscriptstate connected
    set vbArrayAsList [$vb parse -expression "array"]
    if {[llength $vbArrayAsList] != [llength $expectedValues]} {
        error "list size different"
    }
    for {set i 0} {$i < [llength $expectedValues]} {incr i} {
        set vbval [lindex $vbArrayAsList $i]
        set expected [lindex $expectedValues $i]
        if {$vbval ne $expected} {
            set errstring "value does not match (expected:"
            append errstring "\"$expected\", actual: \"$vbval\")"
            error $errstring
        }
    }
    $vb close
    return
}

proc GetStringVarFromVb {} {
    set vb [::activexscripthost::openengine VBScript]
    set TclTestVar "Success"
    $vb setscriptstate connected
    set result [$vb parse -expression {tcl.GetStringVar("TclTestVar")}]
    if {$result ne "Success"} {
        error "unexpected result"
    }
    $vb close
    return
}

} ;# namespace eval
