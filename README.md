# random_stuffs_playground
nothing to read here

###  Alarm

This property gets or sets the behavior of BBAC Capture alarms for all BBAC
Capture channels. Read/Write tlAlarmBehavior.

####  Format

TheHdw.BBACCapture.Alarm(Alarm)

####  Parameters

Alarm

|

Optional. Sets the types of BBAC Capture alarms raised. Type
tlBBACaptureAlarm.  
  
---|---  
  
Alarm is an enumerated constant of type tlBBACaptureAlarm. It can take one of
the following values:

|

tlBBACCaptureAlarmAll

|

Raises alarms for all BBAC Capture alarm types. Default.  
  
---|---  
  
tlBBACCaptureAlarmOverRange

|

Raises alarms only for BBAC Capture over range alarms.  
  
####  Usage

This property is of enumerated type tlAlarmBehavior which has the following
values:

tlAlarmForceFail

|

Forces the test to fail if an alarm of the specified type occurs. Default.  
  
---|---  
  
tlAlarmForceBin

|

Forces the test to fail and bins the part to the error bin if an alarm of the
specified type occurs.  
  
tlAlarmOff

|

Disables the specified alarm in hardware. Removes the time overhead normally
required to process alarms.  
  
tlAlarmDefault

|

Sets the alarm to the default action for the BBAC Capture, which is
tlAlarmForceFail.  
  
tlAlarmContinue

|

Records and reports an alarm of the specified type but does not affect the
pass/fail state or binning. Processes alarms so processing time occurs.  
  
**Caution:**

* * *

Ignored alarms may invalidate test results or harm the instrument.

* * *

####  Example

For all BBACCapture channels, fail any test during which any BBAC Capture
alarm occurs and bin the part to the error bin:

TheHdw.BBACCapture.Alarm(tlBBACCaptureAlarmAll) = tlAlarmForceBin

  

* * *
