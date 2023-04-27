Attribute VB_Name = "VBT_New_Cont"
Option Explicit

Public Function New_Continuity()
On Error GoTo errHandler
    
    Dim RetPinLevels As String
    Dim RetTimeSet As String
    Dim RetEdgeSet As String
    Dim RetDCCategory As String
    Dim RetDCSelector As String
    Dim RetACCategory As String
    Dim RetACSelector As String
    Dim Overlay As String
    
    Call TheExec.DataManager.GetInstanceContext(RetDCCategory, RetDCSelector, RetACCategory, RetACSelector, _
                                                RetTimeSet, RetEdgeSet, RetPinLevels, Overlay)
    
    If RetPinLevels <> TL_C_EMPTYSTR Then thehdw.PinLevels.ApplyPower
    
    
    If RetTimeSet <> TL_C_EMPTYSTR Then Call thehdw.Digital.Timing.Load
    
    Dim CRES_force100ua As New PinListData
    
    Dim DC_Sink_odd As New PinListData
    Dim HSD_Sink_odd As New PinListData
    Dim HVD_Sink_odd As New PinListData
    Dim HVD_Sink_odd_1 As New PinListData
    Dim HVD_Sink_odd_2 As New PinListData
    Dim HVD_Sink_odd_3 As New PinListData
        
    Dim DC_Sink_even As New PinListData
    Dim HSD_Sink_even As New PinListData
    Dim HVD_Sink_even As New PinListData
    Dim HVD_Sink_even_1 As New PinListData
    Dim HVD_Sink_even_2 As New PinListData
    
    Dim DC_odd_vs_VDD As New PinListData
    Dim DC_even_vs_VDD As New PinListData
    Dim HSD_odd_vs_VDD As New PinListData
    Dim HSD_even_vs_VDD As New PinListData
    
    Dim DC_VSSCP_vs_VDDCP As New PinListData
    Dim HSD_M5C1_vs_VDDCP As New PinListData
        
    Dim HSD_PWM_SDI_vs_VDDQ As New PinListData
    Dim HSD_SYNC_SCLK_CS_SDO_vs_VDDQ As New PinListData
    
    Dim DC_SNS1_RCIN_RCTB_TAGB_POUTA_vs_VESDREF As New PinListData
    Dim HSD_HFD_vs_VESDREF As New PinListData
    
    Dim DC_TAGA_RCTA_SRTN_POUTB_vs_VESDREF As New PinListData
    Dim HVD_TEST1_TEST2_vs_VESDREF As New PinListData
    
    Dim DC_M5C2_vs_VSS  As New PinListData
    
    Dim DC_VSSD_vs_VDDD As New PinListData
    
    Dim NEW_CONT As Boolean
        
    Dim SiteNum As Variant

TRIG1

' To connect to the pins directly: K50->CN2A, K17->REF, K6->VDDCP
    SetDatabitsOn ("K50, K6,K16, K17,K4")

' Start continuity_pre

'**************************************************************************************
    
' All DC Pins set to 0v
    thehdw.DCVI.Pins("DC30_CONT_ALL,VSSCP_DC30,VSSD_DC30").Alarm(tlDCVIAlarmMode) = tlAlarmOff
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 3 * ms

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Connect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    'thehdw.Wait 1 * ms
    
' ###########################################################################

    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_I_100ua_2v_M_V_2v").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    thehdw.Wait 3 * ms
    
    CRES_force100ua.AddPin ("CRES_DC30")
    thehdw.DCVI.Pins("CRES_DC30").ClearCaptureMemory

    For Each SiteNum In TheExec.Sites.Active
        If TheExec.Sites.Active Then
             CRES_force100ua = thehdw.DCVI.Pins("CRES_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
        Else
        End If
    Next SiteNum

    Call TestLimitsinflow(CRES_force100ua)
    Call TestLimitsinflow(CRES_force100ua)
   
    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
    End With

    thehdw.DCVI.Pins("SNS1_DC30").Alarm(tlDCVIAlarmAll) = tlAlarmDefault
        
' ###########################################################################

 ' #############
 '   ODD  PINS
 ' #############

    ' Sink 100ua
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_I_N100ua_10v_M_V_10v").Apply
    'thehdw.Wait 2 * ms
    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceI -0.0001
    thehdw.Wait 2 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
       .Mode = tlHVPMUModeCurrent
       .VoltageRange = 2
       .CurrentRange = 100 * uA
       .current = -100 * uA
       .Clamp.Voltage.Max = 3 * v
       .Clamp.Voltage.Min = -3 * v
       .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 2 * ms
     
    thehdw.DCVI.Pins("DC30_CONT_ODD").ClearCaptureMemory
    'thehdw.Wait 1 * ms
       
    Dim TAGA_DC30_OS As New PinListData
    Dim RCTA_DC30_OS As New PinListData
    Dim SRTN_DC30_OS As New PinListData
    Dim VDDD_DC30_OS As New PinListData
    Dim CN1B_DC30_OS As New PinListData
    Dim CN1A_DC30_1_OS As New PinListData
    Dim VLB_DC30_OS As New PinListData
    Dim VDDCP_DC30_OS As New PinListData
    Dim M5C2_DC30_OS As New PinListData
    Dim POUTB_DC30_1_OS As New PinListData
    Dim CN3_DC30_OS As New PinListData
       
    TAGA_DC30_OS = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    RCTA_DC30_OS = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    SRTN_DC30_OS = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDD_DC30_OS = thehdw.DCVI.Pins("VDDD_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1B_DC30_OS = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1A_DC30_1_OS = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)

    VLB_DC30_OS = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDCP_DC30_OS = thehdw.DCVI.Pins("VDDCP_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    M5C2_DC30_OS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    POUTB_DC30_1_OS = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN3_DC30_OS = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    
    Dim PFGB_HSD_OS As New PinListData
    Dim SYNC_HSD_OS As New PinListData
    Dim SDO_HSD_OS As New PinListData
    Dim SDI_HSD_OS As New PinListData
    Dim VDDQ_HSD_OS As New PinListData
    Dim VOUT_HSD_OS As New PinListData
    Dim NCpin35_OS As New PinListData
    Dim NCpin45_OS As New PinListData

    PFGB_HSD_OS = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
    SYNC_HSD_OS = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)
    VDDQ_HSD_OS = thehdw.PPMU.Pins("VDDQ_HSD").Read(tlPPMUReadMeasurements)
    VOUT_HSD_OS = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)
    NCpin35_OS = thehdw.PPMU.Pins("NCpin35").Read(tlPPMUReadMeasurements)
    NCpin45_OS = thehdw.PPMU.Pins("NCpin45").Read(tlPPMUReadMeasurements)

    With thehdw.HVPMU.Pins("HVD_CONT_ODD")
       .Mode = tlHVPMUModeCurrent
       .VoltageRange = 2
       .CurrentRange = 100 * uA
       .current = -100 * uA
       .Clamp.Voltage.Max = 3 * v
       .Clamp.Voltage.Min = -3 * v
       .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

    HVD_Sink_odd_1 = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_odd_2 = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_odd_3 = thehdw.HVPMU.Pins("NCpin23").Read(tlHVPMUVoltage, tlStrobe)
    
    ' Reset
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_V_0v_2ma_M_I_2ma").Apply
    'thehdw.Wait 1 * ms
    
    Call TestLimitsinflow(TAGA_DC30_OS)

    Call TestLimitsinflow(TAGA_DC30_OS)
    
    Call TestLimitsinflow(RCTA_DC30_OS)
    Call TestLimitsinflow(RCTA_DC30_OS)
    
    Call TestLimitsinflow(SRTN_DC30_OS)
    Call TestLimitsinflow(SRTN_DC30_OS)
    
    Call TestLimitsinflow(VDDD_DC30_OS)
    Call TestLimitsinflow(VDDD_DC30_OS)
    
    Call TestLimitsinflow(CN1B_DC30_OS)
    Call TestLimitsinflow(CN1B_DC30_OS)
    
    Call TestLimitsinflow(CN1A_DC30_1_OS)
    Call TestLimitsinflow(CN1A_DC30_1_OS)
    
    Call TestLimitsinflow(VLB_DC30_OS)
    Call TestLimitsinflow(VLB_DC30_OS)
    
    Call TestLimitsinflow(VDDCP_DC30_OS)
    Call TestLimitsinflow(VDDCP_DC30_OS)
    
    Call TestLimitsinflow(M5C2_DC30_OS)
    Call TestLimitsinflow(M5C2_DC30_OS)
    
    Call TestLimitsinflow(POUTB_DC30_1_OS)
    Call TestLimitsinflow(POUTB_DC30_1_OS)
    
    Call TestLimitsinflow(CN3_DC30_OS)
    Call TestLimitsinflow(CN3_DC30_OS)
        
    Call TestLimitsinflow(PFGB_HSD_OS)
    Call TestLimitsinflow(PFGB_HSD_OS)
    Call TestLimitsinflow(SYNC_HSD_OS)
    Call TestLimitsinflow(SYNC_HSD_OS)
    Call TestLimitsinflow(SDO_HSD_OS)
    Call TestLimitsinflow(SDO_HSD_OS)
    Call TestLimitsinflow(SDI_HSD_OS)
    Call TestLimitsinflow(SDI_HSD_OS)
    Call TestLimitsinflow(VDDQ_HSD_OS)
    Call TestLimitsinflow(VDDQ_HSD_OS)
    Call TestLimitsinflow(VOUT_HSD_OS)
    Call TestLimitsinflow(VOUT_HSD_OS)
    Call TestLimitsinflow(NCpin35_OS)
    Call TestLimitsinflow(NCpin35_OS)
    Call TestLimitsinflow(NCpin45_OS)
    Call TestLimitsinflow(NCpin45_OS)
    
    Call TestLimitsinflow(HVD_Sink_odd_1)
    Call TestLimitsinflow(HVD_Sink_odd_1)
    
    Call TestLimitsinflow(HVD_Sink_odd_2)
    Call TestLimitsinflow(HVD_Sink_odd_2)
    
    Call TestLimitsinflow(HVD_Sink_odd_3)
    Call TestLimitsinflow(HVD_Sink_odd_3)

    With thehdw.HVPMU.Pins("HVD_CONT_ODD")
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    
    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceV 0
    thehdw.Wait 1 * ms

' #############
'   EVEN  PINS
' #############

' Sink 100ua
    
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_I_N100ua_10v_M_V_10v").Apply
    thehdw.Wait 1 * ms
    
    Dim SNS1_DC30_OS As New PinListData
    Dim RCIN_DC30_1_OS As New PinListData
    Dim RCTB_DC30_OS As New PinListData
    Dim CN2B_DC30_OS As New PinListData
    Dim CN2A_DC30_OS As New PinListData
    Dim REF_DC30_OS As New PinListData
    Dim VLA_DC30_OS As New PinListData
    Dim VESDREF_DC30_OS As New PinListData
    Dim VDD_DC30_OS As New PinListData
    Dim TAGB_DC30_1_OS As New PinListData
    Dim POUTA_DC30_1_OS As New PinListData

    SNS1_DC30_OS = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe)
    RCIN_DC30_1_OS = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe)
    RCTB_DC30_OS = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe)
    CN2B_DC30_OS = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
    CN2A_DC30_OS = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
    REF_DC30_OS = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)
    VLA_DC30_OS = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
    VESDREF_DC30_OS = thehdw.DCVI.Pins("VESDREF_DC30").Meter.Read(tlStrobe)
    VDD_DC30_OS = thehdw.DCVI.Pins("VDD_DC30").Meter.Read(tlStrobe)
    TAGB_DC30_1_OS = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe)
    POUTA_DC30_1_OS = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe)

    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceI -0.0001
    thehdw.Wait 1 * ms

    Dim PFGA_HSD_OS As New PinListData
    Dim HFD_HSD_OS As New PinListData
    Dim HFG_HSD_OS As New PinListData
    Dim PWM_HSD_OS As New PinListData
    Dim SCLK_HSD_OS As New PinListData
    Dim CS_HSD_OS As New PinListData
    Dim CVF_HSD_OS As New PinListData
    Dim NCpin36_OS As New PinListData
    Dim M5C1_HSD_OS As New PinListData

    PFGA_HSD_OS = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
    HFD_HSD_OS = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements)
    HFG_HSD_OS = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
    PWM_HSD_OS = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
    SCLK_HSD_OS = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
    CS_HSD_OS = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
    CVF_HSD_OS = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)
    NCpin36_OS = thehdw.PPMU.Pins("NCpin36").Read(tlPPMUReadMeasurements)
    M5C1_HSD_OS = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)
    
    With thehdw.HVPMU.Pins("NCpin22,NCpin24")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = -100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 2 * ms

    HVD_Sink_even_1 = thehdw.HVPMU.Pins("NCpin22").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_even_2 = thehdw.HVPMU.Pins("NCpin24").Read(tlHVPMUVoltage, tlStrobe)

' Reset
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_V_0v_2ma_M_I_2ma").Apply
    'thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Disconnect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With

    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceV 0
    
If (0) Then
    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)

    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_2)
    Call TestLimitsinflow(HVD_Sink_even_2)
End If

' **************************************************************************************
    
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
 
' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    'thehdw.Wait 1 * ms
    
'*********************************************************************************************
'*********      continuity on pins considering diode vs VDD ( VDD=0 )       ******************

'**** even

    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("DCcontEVEN_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("HSDcontEVEN_vs_VDD")
        .ForceI 0.0001
        .Connect
    End With
    'thehdw.Wait 1 * ms

If (1) Then
    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)

    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_2)
    Call TestLimitsinflow(HVD_Sink_even_2)
End If

    Dim VLA_DC30_OS_VDD As New PinListData
    Dim CN2B_DC30_OS_VDD As New PinListData
    Dim CN2A_DC30_OS_VDD As New PinListData
    Dim REF_DC30_OS_VDD As New PinListData

    VLA_DC30_OS_VDD = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
    CN2B_DC30_OS_VDD = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
    CN2A_DC30_OS_VDD = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
    REF_DC30_OS_VDD = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)
    
    Dim PFGA_HSD_OS_VDD As New PinListData
    Dim HFG_HSD_OS_VDD As New PinListData
    Dim CVF_HSD_OS_VDD As New PinListData

    PFGA_HSD_OS_VDD = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
    HFG_HSD_OS_VDD = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
    CVF_HSD_OS_VDD = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)

'    Call testlimitsinflow(VLA_DC30_OS_VDD)
'    Call testlimitsinflow(VLA_DC30_OS_VDD)
'    Call testlimitsinflow(CN2B_DC30_OS_VDD)
'    Call testlimitsinflow(CN2B_DC30_OS_VDD)
'    Call testlimitsinflow(CN2A_DC30_OS_VDD)
'    Call testlimitsinflow(CN2A_DC30_OS_VDD)
'    Call testlimitsinflow(REF_DC30_OS_VDD)
'    Call testlimitsinflow(REF_DC30_OS_VDD)
'
'    Call testlimitsinflow(PFGA_HSD_OS_VDD)
'    Call testlimitsinflow(PFGA_HSD_OS_VDD)
'    Call testlimitsinflow(HFG_HSD_OS_VDD)
'    Call testlimitsinflow(HFG_HSD_OS_VDD)
'    Call testlimitsinflow(CVF_HSD_OS_VDD)
'    Call testlimitsinflow(CVF_HSD_OS_VDD)
     
'**** reset

' **************************************************************************************
    
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    'thehdw.Wait 1 * ms

'********************************************************************************
    
    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    
    With thehdw.DCVI.Pins("DCcontODD_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
        .ForceI 0.0001
        .Connect
    End With
    'thehdw.Wait 1 * ms
        
    Call TestLimitsinflow(VLA_DC30_OS_VDD)
    Call TestLimitsinflow(VLA_DC30_OS_VDD)
    Call TestLimitsinflow(CN2B_DC30_OS_VDD)
    Call TestLimitsinflow(CN2B_DC30_OS_VDD)
    Call TestLimitsinflow(CN2A_DC30_OS_VDD)
    Call TestLimitsinflow(CN2A_DC30_OS_VDD)
    Call TestLimitsinflow(REF_DC30_OS_VDD)
    Call TestLimitsinflow(REF_DC30_OS_VDD)
    
    Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
    Call TestLimitsinflow(CVF_HSD_OS_VDD)
    Call TestLimitsinflow(CVF_HSD_OS_VDD)
    
 '***** odd
    Dim CN1B_DC30_OS_VDD As New PinListData
    Dim CN3_DC30_OS_VDD As New PinListData
    Dim CN1A_DC30_1_OS_VDD As New PinListData
    Dim VLB_DC30_OS_VDD As New PinListData

    CN1B_DC30_OS_VDD = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe)
    CN3_DC30_OS_VDD = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe)
    CN1A_DC30_1_OS_VDD = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe)
    VLB_DC30_OS_VDD = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe)
        
    Dim PFGB_HSD_OS_VDD As New PinListData
    Dim VOUT_HSD_OS_VDD As New PinListData
    
    PFGB_HSD_OS_VDD = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
    VOUT_HSD_OS_VDD = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.DCVI.Pins("DCcontODD_vs_VDD")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Disconnect
        .Gate = False
    End With

    With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
        .ForceI 0.0001
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
    
     With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With

'    Call testlimitsinflow(CN1B_DC30_OS_VDD)
'    Call testlimitsinflow(CN1B_DC30_OS_VDD)
'    Call testlimitsinflow(CN3_DC30_OS_VDD)
'    Call testlimitsinflow(CN3_DC30_OS_VDD)
'    Call testlimitsinflow(CN1A_DC30_1_OS_VDD)
'    Call testlimitsinflow(CN1A_DC30_1_OS_VDD)
'    Call testlimitsinflow(VLB_DC30_OS_VDD)
'    Call testlimitsinflow(VLB_DC30_OS_VDD)
'
'    Call testlimitsinflow(PFGB_HSD_OS_VDD)
'    Call testlimitsinflow(PFGB_HSD_OS_VDD)
'    Call testlimitsinflow(VOUT_HSD_OS_VDD)
'    Call testlimitsinflow(VOUT_HSD_OS_VDD)

'********************************************************************************************
'********************************************************************************************

'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDCP ( VDDCP=0 )       ******************
        
    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 3 * ms
    
    With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms

   With thehdw.PPMU.Pins("M5C1_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms
    

    thehdw.DCVI.Pins("VSSCP_DC30").ClearCaptureMemory
            
    DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe)
    DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)

' Instable on site0 during GM debug also with OLD rev; HW issue???
'    For Each SiteNum In TheExec.Sites.Active
'           If DC_VSSCP_vs_VDDCP < 300 * mV Then
'            thehdw.DCVI.Pins("VSSCP_DC30").ClearCaptureMemory
'
'            DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe)
'         End If
'    Next SiteNum
'        thehdw.Wait 3 * ms
'    For Each SiteNum In TheExec.Sites.Active
'           If DC_VSSCP_vs_VDDCP < 300 * mV Then
'            thehdw.DCVI.Pins("VSSCP_DC30").ClearCaptureMemory
'
'            DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
'         End If
'    Next SiteNum
' Instable on site0 during GM debug also with OLD rev; HW issue???

    HSD_M5C1_vs_VDDCP = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)
    
    Call TestLimitsinflow(CN1B_DC30_OS_VDD)
    Call TestLimitsinflow(CN1B_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
    Call TestLimitsinflow(VLB_DC30_OS_VDD)
    Call TestLimitsinflow(VLB_DC30_OS_VDD)
  
    Call TestLimitsinflow(PFGB_HSD_OS_VDD)
    Call TestLimitsinflow(PFGB_HSD_OS_VDD)
    Call TestLimitsinflow(VOUT_HSD_OS_VDD)
    Call TestLimitsinflow(VOUT_HSD_OS_VDD)
    


    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    
    With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
    thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("M5C1_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
    thehdw.Wait 1 * ms

    Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
    Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)
   
'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDQ ( VDDQ=0 )       **************

    With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms
    
'***** even

    Dim PWM_HSD_OS_VDDQ As New PinListData
    Dim SDI_HSD_OS_VDDQ As New PinListData

    PWM_HSD_OS_VDDQ = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
'****** odd

    With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms

    Dim SYNC_HSD_OS_VDDQ As New PinListData
    Dim SCLK_HSD_OS_VDDQ As New PinListData
    Dim CS_HSD_OS_VDDQ As New PinListData
    Dim SDO_HSD_OS_VDDQ As New PinListData

 
    SYNC_HSD_OS_VDDQ = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SCLK_HSD_OS_VDDQ = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
    CS_HSD_OS_VDDQ = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
    
    With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

    Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
    Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
    
    Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
    Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
    Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
    Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
    Call TestLimitsinflow(CS_HSD_OS_VDDQ)
    Call TestLimitsinflow(CS_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDO_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDO_HSD_OS_VDDQ)

'*********************************************************************************************
'*********      continuity on pins considering diode vs VESDREF ( VESDREF=0 )       **************


    With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 75 * v
        .CurrentRange = 32 * mA
        .Clamp.current.Max = 32 * mA
        .Clamp.current.Min = -32 * mA
        .Voltage = 0 * v
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    thehdw.Wait 2 * ms

    With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms
 
    With thehdw.PPMU.Pins("HFD_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 2 * ms
    
'**** even
    
    Dim SNS1_DC30_OS_VESDREF As New PinListData
    Dim RCIN_DC30_1_OS_VESDREF As New PinListData
    Dim RCTB_DC30_OS_VESDREF As New PinListData
    Dim TAGB_DC30_1_OS_VESDREF As New PinListData
    Dim POUTA_DC30_1_OS_VESDREF As New PinListData
    
    SNS1_DC30_OS_VESDREF = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCIN_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    TAGB_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    POUTA_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)

    HSD_HFD_vs_VESDREF = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements, 100)

    With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Gate = False
        .Disconnect
    End With
    
    With thehdw.PPMU.Pins("HFD_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    
    Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
    Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
    Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
    Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
    Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
    Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
    Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
    Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
    
    Call TestLimitsinflow(HSD_HFD_vs_VESDREF)
    Call TestLimitsinflow(HSD_HFD_vs_VESDREF)


'**************************************************************************************
'**************************************************************************************

'********************************************************************************************
'********************************************************************************************

'    With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
'        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
'        .Gate = True
'        .Connect
'    End With
'    thehdw.Wait 1 * ms

'**** odd
    thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
    thehdw.Wait 2 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
         .Mode = tlHVPMUModeCurrent
         .VoltageRange = 2
         .CurrentRange = 100 * uA
         .current = 100 * uA
         .Clamp.Voltage.Max = 3 * v
         .Clamp.Voltage.Min = -3 * v
         .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 4 * ms
    
    Dim TAGA_DC30_OS_VESDREF As New PinListData
    Dim RCTA_DC30_OS_VESDREF As New PinListData
    Dim SRTN_DC30_OS_VESDREF As New PinListData
    Dim POUTB_DC30_OS_VESDREF As New PinListData

    TAGA_DC30_OS_VESDREF = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe)
    RCTA_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe)
    SRTN_DC30_OS_VESDREF = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe)
    POUTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe)
    
    Dim TEST1_HSDHV_OS_VESDREF As New PinListData
    Dim TEST2_HSDHV_OS_VESDREF As New PinListData

    TEST1_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    TEST2_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)

    thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
    'thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23,NCpin22,NCpin24")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    'thehdw.Wait 1 * ms

'    Call testlimitsinflow(TAGA_DC30_OS_VESDREF)
'    Call testlimitsinflow(TAGA_DC30_OS_VESDREF)
'    Call testlimitsinflow(RCTA_DC30_OS_VESDREF)
'    Call testlimitsinflow(RCTA_DC30_OS_VESDREF)
'    Call testlimitsinflow(SRTN_DC30_OS_VESDREF)
'    Call testlimitsinflow(SRTN_DC30_OS_VESDREF)
'    Call testlimitsinflow(POUTB_DC30_OS_VESDREF)
'    Call testlimitsinflow(POUTB_DC30_OS_VESDREF)
'
'    Call testlimitsinflow(TEST1_HSDHV_OS_VESDREF)
'    Call testlimitsinflow(TEST1_HSDHV_OS_VESDREF)
'    Call testlimitsinflow(TEST2_HSDHV_OS_VESDREF)
'    Call testlimitsinflow(TEST2_HSDHV_OS_VESDREF)

    With thehdw.DCVI.Pins("VESDREF_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = False
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

    Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
    Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)

    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)

'**************************************************************************************
'**************************************************************************************


'********************************************************************************************
'********************************************************************************************
'*********************************************************************************************
'*********      continuity on pins M5C2 considering diode vs VSS ( VSS=0 )       **************

    With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

    DC_M5C2_vs_VSS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe)

    With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
    thehdw.Wait 1 * ms
        .Gate = False
        .Disconnect
    End With
    thehdw.Wait 1 * ms
  
    Call TestLimitsinflow(DC_M5C2_vs_VSS)
    Call TestLimitsinflow(DC_M5C2_vs_VSS)
 
'*********************************************************************************************
'*********************************************************************************************
'**************************************************************************************
'**************************************************************************************


'*********************************************************************************************
'*********      continuity on pins VSSD considering diode vs VDDD ( VDDD=0 )       ******************

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
    DC_VSSD_vs_VDDD = thehdw.DCVI.Pins("VSSD_DC30").Meter.Read(tlStrobe)

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
    thehdw.Wait 1 * ms
     
    Call TestLimitsinflow(DC_VSSD_vs_VDDD)
    Call TestLimitsinflow(DC_VSSD_vs_VDDD)

'*********************************************************************************************

'disconnect
    thehdw.DCVI.Pins("DC30_CONT_ALL").Disconnect (tlDCVIConnectDefault)
    thehdw.DCVI.Pins("DC30_CONT_ALL,VSSCP_DC30,VSSD_DC30").Alarm(tlDCVIAlarmMode) = tlAlarmDefault

    thehdw.PPMU.Pins("HSD_CONT_ALL").Disconnect
    
     With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
        .Gate = False
        .Disconnect
    End With
    thehdw.Wait 1 * ms

    SetDatabitsOff ("K50, K6,K16, K17,K4")

    With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = False
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    
    thehdw.Digital.DisconnectPins ("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
    thehdw.Wait 1 * ms
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function New_Continuity_ori()
On Error GoTo errHandler

Dim RetPinLevels As String
Dim RetTimeSet As String
Dim RetEdgeSet As String
Dim RetDCCategory As String
Dim RetDCSelector As String
Dim RetACCategory As String
Dim RetACSelector As String
Dim Overlay As String

Call TheExec.DataManager.GetInstanceContext(RetDCCategory, RetDCSelector, RetACCategory, RetACSelector, _
                                            RetTimeSet, RetEdgeSet, RetPinLevels, Overlay)

If RetPinLevels <> TL_C_EMPTYSTR Then thehdw.PinLevels.ApplyPower


If RetTimeSet <> TL_C_EMPTYSTR Then Call thehdw.Digital.Timing.Load

Dim CRES_force100ua As New PinListData

Dim DC_Sink_odd As New PinListData
Dim HSD_Sink_odd As New PinListData
Dim HVD_Sink_odd As New PinListData
Dim HVD_Sink_odd_1 As New PinListData
Dim HVD_Sink_odd_2 As New PinListData
Dim HVD_Sink_odd_3 As New PinListData
    
Dim DC_Sink_even As New PinListData
Dim HSD_Sink_even As New PinListData
Dim HVD_Sink_even As New PinListData
Dim HVD_Sink_even_1 As New PinListData
Dim HVD_Sink_even_2 As New PinListData

Dim DC_odd_vs_VDD As New PinListData
Dim DC_even_vs_VDD As New PinListData
Dim HSD_odd_vs_VDD As New PinListData
Dim HSD_even_vs_VDD As New PinListData

Dim DC_VSSCP_vs_VDDCP As New PinListData
Dim HSD_M5C1_vs_VDDCP As New PinListData
    
Dim HSD_PWM_SDI_vs_VDDQ As New PinListData
Dim HSD_SYNC_SCLK_CS_SDO_vs_VDDQ As New PinListData

Dim DC_SNS1_RCIN_RCTB_TAGB_POUTA_vs_VESDREF As New PinListData
Dim HSD_HFD_vs_VESDREF As New PinListData

Dim DC_TAGA_RCTA_SRTN_POUTB_vs_VESDREF As New PinListData
Dim HVD_TEST1_TEST2_vs_VESDREF As New PinListData

Dim DC_M5C2_vs_VSS  As New PinListData

Dim DC_VSSD_vs_VDDD As New PinListData

Dim NEW_CONT As Boolean
    
Dim SiteNum As Variant
    
'If RetPinLevels <> TL_C_EMPTYSTR Then thehdw.PinLevels.ApplyPower

' To connect to the pins directly: K50->CN2A, K17->REF, K6->VDDCP
    SetDatabitsOn ("K50, K6,K16, K17,K4")
'    SetDatabitsOn ("K16, K17")
'    SetDatabitsOn ("K4") '**** connect VSSD and VSSCP to dc30; it may be considered as normal pins ( testing continuity )
'    thehdw.Wait 1 * ms
    
' Start continuity_pre

'**************************************************************************************
    
' All DC Pins set to 0v
    thehdw.DCVI.Pins("DC30_CONT_ALL,VSSCP_DC30,VSSD_DC30").Alarm(tlDCVIAlarmMode) = tlAlarmOff
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
'
'      With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
'        .PSets("f_V_0v_2ma_M_I_2ma").Apply
'        .Gate = True
'        .Connect
'    End With
'     thehdw.Wait 1 * ms
     
      
' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
''        .current = 2 * mA
        .ForceV 0
        .Connect
    End With
    thehdw.Wait 1 * ms
    
   
     
    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
    
' ###########################################################################

    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_I_100ua_2v_M_V_2v").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms
       
 CRES_force100ua.AddPin ("CRES_DC30")
thehdw.DCVI.Pins("CRES_DC30").ClearCaptureMemory

    For Each SiteNum In TheExec.Sites.Active
           If TheExec.Sites.Active Then
             CRES_force100ua = thehdw.DCVI.Pins("CRES_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
         Else
         End If
    Next SiteNum

    
    Call TestLimitsinflow(CRES_force100ua)
       Call TestLimitsinflow(CRES_force100ua)
   
    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
    End With
'    thehdw.Wait 1 * ms
    
    
  thehdw.DCVI.Pins("SNS1_DC30").Alarm(tlDCVIAlarmAll) = tlAlarmDefault
    
    
' ###########################################################################

 ' #############
 '   ODD  PINS
 ' #############

    ' Sink 100ua
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_I_N100ua_10v_M_V_10v").Apply
    thehdw.Wait 2 * ms

    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceI -0.0001
    thehdw.Wait 2 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
       .Mode = tlHVPMUModeCurrent
       .VoltageRange = 2
       .CurrentRange = 100 * uA
       .current = -100 * uA
       .Clamp.Voltage.Max = 3 * v
       .Clamp.Voltage.Min = -3 * v
       .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 2 * ms
     
    thehdw.DCVI.Pins("DC30_CONT_ODD").ClearCaptureMemory
    thehdw.Wait 1 * ms
   
''    DC_Sink_odd = TheHdw.DCVI.Pins("DC30_CONT_ODD").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
       
    Dim TAGA_DC30_OS As New PinListData
    Dim RCTA_DC30_OS As New PinListData
    Dim SRTN_DC30_OS As New PinListData
    Dim VDDD_DC30_OS As New PinListData
    Dim CN1B_DC30_OS As New PinListData
    Dim CN1A_DC30_1_OS As New PinListData
    Dim VLB_DC30_OS As New PinListData
    Dim VDDCP_DC30_OS As New PinListData
    Dim M5C2_DC30_OS As New PinListData
    Dim POUTB_DC30_1_OS As New PinListData
    Dim CN3_DC30_OS As New PinListData
       
    TAGA_DC30_OS = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    RCTA_DC30_OS = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    SRTN_DC30_OS = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDD_DC30_OS = thehdw.DCVI.Pins("VDDD_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1B_DC30_OS = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1A_DC30_1_OS = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)

    VLB_DC30_OS = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDCP_DC30_OS = thehdw.DCVI.Pins("VDDCP_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    M5C2_DC30_OS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    POUTB_DC30_1_OS = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN3_DC30_OS = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    
    Dim PFGB_HSD_OS As New PinListData
    Dim SYNC_HSD_OS As New PinListData
    Dim SDO_HSD_OS As New PinListData
    Dim SDI_HSD_OS As New PinListData
    Dim VDDQ_HSD_OS As New PinListData
    Dim VOUT_HSD_OS As New PinListData
    Dim NCpin35_OS As New PinListData
    Dim NCpin45_OS As New PinListData

    PFGB_HSD_OS = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
    SYNC_HSD_OS = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)
    VDDQ_HSD_OS = thehdw.PPMU.Pins("VDDQ_HSD").Read(tlPPMUReadMeasurements)
    VOUT_HSD_OS = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)
    NCpin35_OS = thehdw.PPMU.Pins("NCpin35").Read(tlPPMUReadMeasurements)
    NCpin45_OS = thehdw.PPMU.Pins("NCpin45").Read(tlPPMUReadMeasurements)


    With thehdw.HVPMU.Pins("HVD_CONT_ODD")
       .Mode = tlHVPMUModeCurrent
       .VoltageRange = 2
       .CurrentRange = 100 * uA
       .current = -100 * uA
       .Clamp.Voltage.Max = 3 * v
       .Clamp.Voltage.Min = -3 * v
       .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
   
'''   HVD_Sink_odd = thehdw.HVPMU.Pins("HVD_CONT_ODD").Read(tlHVPMUVoltage, tlStrobe)
   HVD_Sink_odd_1 = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_odd_2 = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
   HVD_Sink_odd_3 = thehdw.HVPMU.Pins("NCpin23").Read(tlHVPMUVoltage, tlStrobe)
 
''    HVD_Sink_odd_1.AddPin ("TEST1_HSDHV")
''    HVD_Sink_odd_2.AddPin ("TEST2_HSDHV")
''    HVD_Sink_odd_3.AddPin ("NCpin23")
''
''    For Each SiteNum In TheExec.Sites.Active
''         If TheExec.TesterMode = testModeOnline Then
''              HVD_Sink_odd_1.Pins("TEST1_HSDHV").value(SiteNum) = -0.6
''              HVD_Sink_odd_2.Pins("TEST2_HSDHV").value(SiteNum) = -0.6
''              HVD_Sink_odd_3.Pins("NCpin23").value(SiteNum) = -3
''         Else
''         End If
''    Next SiteNum

   ' Reset
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_V_0v_2ma_M_I_2ma").Apply
    thehdw.Wait 1 * ms
    
'''     Call testlimitsinflow(DC_Sink_odd)
      Call TestLimitsinflow(TAGA_DC30_OS)
       Call TestLimitsinflow(TAGA_DC30_OS)

       Call TestLimitsinflow(RCTA_DC30_OS)
       Call TestLimitsinflow(RCTA_DC30_OS)
       
      Call TestLimitsinflow(SRTN_DC30_OS)
       Call TestLimitsinflow(SRTN_DC30_OS)
       
      Call TestLimitsinflow(VDDD_DC30_OS)
       Call TestLimitsinflow(VDDD_DC30_OS)
       
      Call TestLimitsinflow(CN1B_DC30_OS)
       Call TestLimitsinflow(CN1B_DC30_OS)
       
      Call TestLimitsinflow(CN1A_DC30_1_OS)
       Call TestLimitsinflow(CN1A_DC30_1_OS)
       
      Call TestLimitsinflow(VLB_DC30_OS)
       Call TestLimitsinflow(VLB_DC30_OS)
       
       Call TestLimitsinflow(VDDCP_DC30_OS)
       Call TestLimitsinflow(VDDCP_DC30_OS)
       
        Call TestLimitsinflow(M5C2_DC30_OS)
       Call TestLimitsinflow(M5C2_DC30_OS)
       
        Call TestLimitsinflow(POUTB_DC30_1_OS)
       Call TestLimitsinflow(POUTB_DC30_1_OS)

        Call TestLimitsinflow(CN3_DC30_OS)
       Call TestLimitsinflow(CN3_DC30_OS)


Call TestLimitsinflow(PFGB_HSD_OS)
Call TestLimitsinflow(PFGB_HSD_OS)
Call TestLimitsinflow(SYNC_HSD_OS)
Call TestLimitsinflow(SYNC_HSD_OS)
Call TestLimitsinflow(SDO_HSD_OS)
Call TestLimitsinflow(SDO_HSD_OS)
Call TestLimitsinflow(SDI_HSD_OS)
Call TestLimitsinflow(SDI_HSD_OS)
Call TestLimitsinflow(VDDQ_HSD_OS)
Call TestLimitsinflow(VDDQ_HSD_OS)
Call TestLimitsinflow(VOUT_HSD_OS)
Call TestLimitsinflow(VOUT_HSD_OS)
Call TestLimitsinflow(NCpin35_OS)
Call TestLimitsinflow(NCpin35_OS)
Call TestLimitsinflow(NCpin45_OS)
Call TestLimitsinflow(NCpin45_OS)



''    Call testlimitsinflow(HSD_Sink_odd)
    
    Call TestLimitsinflow(HVD_Sink_odd_1)
        Call TestLimitsinflow(HVD_Sink_odd_1)

    Call TestLimitsinflow(HVD_Sink_odd_2)
        Call TestLimitsinflow(HVD_Sink_odd_2)

    Call TestLimitsinflow(HVD_Sink_odd_3)
    Call TestLimitsinflow(HVD_Sink_odd_3)
  
   
  With thehdw.HVPMU.Pins("HVD_CONT_ODD")
'''        .current = 2 * mA
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    
    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceV 0
 thehdw.Wait 1 * ms

' #############
 '   EVEN  PINS
 ' #############

    ' Sink 100ua
    
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_I_N100ua_10v_M_V_10v").Apply
     thehdw.Wait 1 * ms
   
Dim SNS1_DC30_OS As New PinListData
Dim RCIN_DC30_1_OS As New PinListData
Dim RCTB_DC30_OS As New PinListData
Dim CN2B_DC30_OS As New PinListData
Dim CN2A_DC30_OS As New PinListData
Dim REF_DC30_OS As New PinListData
Dim VLA_DC30_OS As New PinListData
Dim VESDREF_DC30_OS As New PinListData
Dim VDD_DC30_OS As New PinListData
Dim TAGB_DC30_1_OS As New PinListData
Dim POUTA_DC30_1_OS As New PinListData



    SNS1_DC30_OS = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe)
    RCIN_DC30_1_OS = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe)
    RCTB_DC30_OS = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe)
    CN2B_DC30_OS = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
    CN2A_DC30_OS = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
    REF_DC30_OS = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)
    VLA_DC30_OS = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
    VESDREF_DC30_OS = thehdw.DCVI.Pins("VESDREF_DC30").Meter.Read(tlStrobe)
    VDD_DC30_OS = thehdw.DCVI.Pins("VDD_DC30").Meter.Read(tlStrobe)
    TAGB_DC30_1_OS = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe)
    POUTA_DC30_1_OS = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe)

'''    DC_Sink_even = TheHdw.DCVI.Pins("DC30_CONT_EVEN").Meter.Read(tlStrobe)
  
    
    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceI -0.0001
    thehdw.Wait 1 * ms


Dim PFGA_HSD_OS As New PinListData
Dim HFD_HSD_OS As New PinListData
Dim HFG_HSD_OS As New PinListData
Dim PWM_HSD_OS As New PinListData
Dim SCLK_HSD_OS As New PinListData
Dim CS_HSD_OS As New PinListData
Dim CVF_HSD_OS As New PinListData
Dim NCpin36_OS As New PinListData
Dim M5C1_HSD_OS As New PinListData

   PFGA_HSD_OS = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
   HFD_HSD_OS = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements)
   HFG_HSD_OS = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
   PWM_HSD_OS = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
   SCLK_HSD_OS = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
   CS_HSD_OS = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
   CVF_HSD_OS = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)
   NCpin36_OS = thehdw.PPMU.Pins("NCpin36").Read(tlPPMUReadMeasurements)
   M5C1_HSD_OS = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)



'''   HSD_Sink_even = TheHdw.PPMU.Pins("HSD_CONT_EVEN").Read(tlPPMUReadMeasurements)
    
   With thehdw.HVPMU.Pins("NCpin22,NCpin24")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = -100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
   End With
thehdw.Wait 2 * ms

  HVD_Sink_even_1 = thehdw.HVPMU.Pins("NCpin22").Read(tlHVPMUVoltage, tlStrobe)
   HVD_Sink_even_2 = thehdw.HVPMU.Pins("NCpin24").Read(tlHVPMUVoltage, tlStrobe)
  
'''    HVD_Sink_even_1.AddPin ("TEST1_HSDHV")
'''    HVD_Sink_even_2.AddPin ("TEST2_HSDHV")
'''
'''    For Each SiteNum In TheExec.Sites.Active
'''         If TheExec.TesterMode = testModeOnline Then
'''              HVD_Sink_even_1.Pins("TEST1_HSDHV").value(SiteNum) = -3
'''              HVD_Sink_even_2.Pins("TEST2_HSDHV").value(SiteNum) = -3
'''         Else
'''         End If
'''    Next SiteNum



' Reset
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_V_0v_2ma_M_I_2ma").Apply
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Disconnect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With

    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceV 0
    
   
    Call TestLimitsinflow(SNS1_DC30_OS)
     Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
      Call TestLimitsinflow(RCIN_DC30_1_OS)
   Call TestLimitsinflow(RCTB_DC30_OS)
     Call TestLimitsinflow(RCTB_DC30_OS)
   Call TestLimitsinflow(CN2B_DC30_OS)
     Call TestLimitsinflow(CN2B_DC30_OS)
  Call TestLimitsinflow(CN2A_DC30_OS)
   Call TestLimitsinflow(CN2A_DC30_OS)
   Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
      Call TestLimitsinflow(VLA_DC30_OS)
   Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
     Call TestLimitsinflow(VDD_DC30_OS)
      Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
      Call TestLimitsinflow(TAGB_DC30_1_OS)
   Call TestLimitsinflow(POUTA_DC30_1_OS)
   Call TestLimitsinflow(POUTA_DC30_1_OS)

'''     Call testlimitsinflow(DC_Sink_even)
     Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
   Call TestLimitsinflow(SCLK_HSD_OS)
       Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(NCpin36_OS)
   Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
   Call TestLimitsinflow(M5C1_HSD_OS)
    
    
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_1)
        Call TestLimitsinflow(HVD_Sink_even_2)
    Call TestLimitsinflow(HVD_Sink_even_2)

   
    
' **************************************************************************************
    


' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms
     
'      With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
'        .PSets("f_V_0v_2ma_M_I_2ma").Apply
'        .Gate = True
'        .Disconnect
'    End With
'     thehdw.Wait 1 * ms
     
      
' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
''        .current = 2 * mA
        .ForceV 0
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    
   
' All HSD pins at 0V, using HVPPMU
'''    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
'''        .Voltage = 0
'''        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
'''    End With
     thehdw.Wait 1 * ms
    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
    
'*********************************************************************************************
'*********      continuity on pins considering diode vs VDD ( VDD=0 )       ******************

'**** even

    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
    
  With thehdw.DCVI.Pins("DCcontEVEN_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
 With thehdw.PPMU.Pins("HSDcontEVEN_vs_VDD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms
    
'''    thehdw.PPMU.Pins("HSDcontEVEN_vs_VDD").ForceI 0.0001




Dim VLA_DC30_OS_VDD As New PinListData
Dim CN2B_DC30_OS_VDD As New PinListData
Dim CN2A_DC30_OS_VDD As New PinListData
Dim REF_DC30_OS_VDD As New PinListData



   VLA_DC30_OS_VDD = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
   CN2B_DC30_OS_VDD = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
   CN2A_DC30_OS_VDD = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
   REF_DC30_OS_VDD = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)

'''   DC_even_vs_VDD = TheHdw.DCVI.Pins("DCcontEVEN_vs_VDD").Meter.Read(tlStrobe)

Dim PFGA_HSD_OS_VDD As New PinListData
Dim HFG_HSD_OS_VDD As New PinListData
Dim CVF_HSD_OS_VDD As New PinListData




   PFGA_HSD_OS_VDD = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
   HFG_HSD_OS_VDD = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
   CVF_HSD_OS_VDD = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)

'''   HSD_even_vs_VDD = TheHdw.PPMU.Pins("HSDcontEVEN_vs_VDD").Read(tlPPMUReadMeasurements)


    Call TestLimitsinflow(VLA_DC30_OS_VDD)
     Call TestLimitsinflow(VLA_DC30_OS_VDD)
         Call TestLimitsinflow(CN2B_DC30_OS_VDD)
     Call TestLimitsinflow(CN2B_DC30_OS_VDD)
    Call TestLimitsinflow(CN2A_DC30_OS_VDD)
     Call TestLimitsinflow(CN2A_DC30_OS_VDD)
    Call TestLimitsinflow(REF_DC30_OS_VDD)
     Call TestLimitsinflow(REF_DC30_OS_VDD)

   Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
   Call TestLimitsinflow(CVF_HSD_OS_VDD)
    Call TestLimitsinflow(CVF_HSD_OS_VDD)
  
   
   '**** reset

' **************************************************************************************
    
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms
     
'      With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
'        .PSets("f_V_0v_2ma_M_I_2ma").Apply
'        .Gate = True
'        .Disconnect
'    End With
'     thehdw.Wait 1 * ms
     
      
' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
''        .current = 2 * mA
        .ForceV 0
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    
   
' All HSD pins at 0V, using HVPPMU
'''    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
'''        .Voltage = 0
'''        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
'''    End With
     thehdw.Wait 1 * ms
    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

'********************************************************************************
    
       With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
'     thehdw.Wait 1 * ms
    
  With thehdw.DCVI.Pins("DCcontODD_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
 With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms
    
 '***** odd

Dim CN1B_DC30_OS_VDD As New PinListData
Dim CN3_DC30_OS_VDD As New PinListData
Dim CN1A_DC30_1_OS_VDD As New PinListData
Dim VLB_DC30_OS_VDD As New PinListData



    CN1B_DC30_OS_VDD = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe)
    CN3_DC30_OS_VDD = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe)
    CN1A_DC30_1_OS_VDD = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe)
    VLB_DC30_OS_VDD = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe)


Dim PFGB_HSD_OS_VDD As New PinListData
Dim VOUT_HSD_OS_VDD As New PinListData


''    DC_odd_vs_VDD = TheHdw.DCVI.Pins("DCcontODD_vs_VDD").Meter.Read(tlStrobe)
   

   PFGB_HSD_OS_VDD = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
   VOUT_HSD_OS_VDD = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)

    
    With thehdw.DCVI.Pins("DCcontODD_vs_VDD")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Disconnect
        .Gate = False
    End With
    

    With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    
     With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms
     
     
    Call TestLimitsinflow(CN1B_DC30_OS_VDD)
     Call TestLimitsinflow(CN1B_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
     Call TestLimitsinflow(VLB_DC30_OS_VDD)
    Call TestLimitsinflow(VLB_DC30_OS_VDD)
  
   Call TestLimitsinflow(PFGB_HSD_OS_VDD)
   Call TestLimitsinflow(PFGB_HSD_OS_VDD)
   Call TestLimitsinflow(VOUT_HSD_OS_VDD)
   Call TestLimitsinflow(VOUT_HSD_OS_VDD)

   
   
'********************************************************************************************
'********************************************************************************************


'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDCP ( VDDCP=0 )       ******************
    
    
    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
    
  With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
 With thehdw.PPMU.Pins("M5C1_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms
    
'''    thehdw.DCVI.Pins("VSSCP_DC30").PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
'''    thehdw.Wait 1 * ms
'''
'''    thehdw.PPMU.Pins("M5C1_HSD").ForceI 0.0001
'''    thehdw.Wait 1 * ms

    DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe, 100, 100000)

    HSD_M5C1_vs_VDDCP = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms
    
  With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms
     
 With thehdw.PPMU.Pins("M5C1_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Disconnect
    End With
    thehdw.Wait 1 * ms

 Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
 Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)

  
   
'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDQ ( VDDQ=0 )       **************


    With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

     
 With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms
    
'***** even

 Dim PWM_HSD_OS_VDDQ As New PinListData
  Dim SDI_HSD_OS_VDDQ As New PinListData

 
    PWM_HSD_OS_VDDQ = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)

     With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Disconnect
    End With
'    thehdw.Wait 1 * ms
'****** odd

 With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 1 * ms

 Dim SYNC_HSD_OS_VDDQ As New PinListData
  Dim SCLK_HSD_OS_VDDQ As New PinListData
 Dim CS_HSD_OS_VDDQ As New PinListData
  Dim SDO_HSD_OS_VDDQ As New PinListData

 
    SYNC_HSD_OS_VDDQ = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SCLK_HSD_OS_VDDQ = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
    CS_HSD_OS_VDDQ = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)

'''    HSD_SYNC_SCLK_CS_SDO_vs_VDDQ = TheHdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Disconnect
    End With
'    thehdw.Wait 1 * ms
    
     With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms
    
     Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
      Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
     Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
     Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
   
   Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
   Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
   Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
   Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
   Call TestLimitsinflow(CS_HSD_OS_VDDQ)
   Call TestLimitsinflow(CS_HSD_OS_VDDQ)
   Call TestLimitsinflow(SDO_HSD_OS_VDDQ)
   Call TestLimitsinflow(SDO_HSD_OS_VDDQ)

  
   
'*********************************************************************************************
'*********      continuity on pins considering diode vs VESDREF ( VESDREF=0 )       **************


  With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 75 * v
        .CurrentRange = 32 * mA
        .Clamp.current.Max = 32 * mA
        .Clamp.current.Min = -32 * mA
        .Voltage = 0 * v
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    thehdw.Wait 2 * ms

    With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
     
 With thehdw.PPMU.Pins("HFD_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 2 * ms
    
'**** even
    
 Dim SNS1_DC30_OS_VESDREF As New PinListData
 Dim RCIN_DC30_1_OS_VESDREF As New PinListData
 Dim RCTB_DC30_OS_VESDREF As New PinListData
 Dim TAGB_DC30_1_OS_VESDREF As New PinListData
 Dim POUTA_DC30_1_OS_VESDREF As New PinListData
    
    SNS1_DC30_OS_VESDREF = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCIN_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    TAGB_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    POUTA_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)


   HSD_HFD_vs_VESDREF = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements, 100)


 With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Gate = False
        .Disconnect
    End With
'    thehdw.Wait 1 * ms
    
    With thehdw.PPMU.Pins("HFD_HSD")
''        .current = 2 * mA
        .ForceI 0.0001
        .Disconnect
    End With
    thehdw.Wait 1 * ms

  Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
   Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
  Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
  Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
  Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
   
   Call TestLimitsinflow(HSD_HFD_vs_VESDREF)
   Call TestLimitsinflow(HSD_HFD_vs_VESDREF)


'**************************************************************************************
'**************************************************************************************

'********************************************************************************************
'********************************************************************************************


 With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
     
'**** odd
    thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
    thehdw.Wait 2 * ms

    
   With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = 100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
   End With
     thehdw.Wait 4 * ms
     
'''    With TheHdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV")
'''        .Mode = tlHVPMUModeCurrent
'''        .VoltageRange = 2
'''        .CurrentRange = 100 * uA
'''        .current = 100 * uA
'''        .Clamp.Voltage.Max = 2 * v
'''        .Clamp.Voltage.Min = -2 * v
'''        .Connect (tlHVPMUConnectForce)
'''    End With
'''    TheHdw.Wait 5 * ms
    
Dim TAGA_DC30_OS_VESDREF As New PinListData
Dim RCTA_DC30_OS_VESDREF As New PinListData
Dim SRTN_DC30_OS_VESDREF As New PinListData
Dim POUTB_DC30_OS_VESDREF As New PinListData


    TAGA_DC30_OS_VESDREF = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe)
    RCTA_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe)
    SRTN_DC30_OS_VESDREF = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe)
    POUTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe)


'''    HVD_TEST1_vs_VESDREF.AddPin ("TEST1_HSDHV")
'''    HVD_TEST2_vs_VESDREF.AddPin ("TEST2_HSDHV")
'''
'''    For Each SiteNum In TheExec.Sites.Active
'''         If TheExec.TesterMode = testModeOnline Then
'''              HVD_TEST1_vs_VESDREF.Pins("TEST1_HSDHV").value(SiteNum) = 0.6
'''              HVD_TEST2_vs_VESDREF.Pins("TEST2_HSDHV").value(SiteNum) = 0.6
'''         Else
'''         End If
'''    Next SiteNum
    
Dim TEST1_HSDHV_OS_VESDREF As New PinListData
Dim TEST2_HSDHV_OS_VESDREF As New PinListData

    TEST1_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    TEST2_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)

   thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23,NCpin22,NCpin24")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
thehdw.Wait 1 * ms

    Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
     Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)
   
    
    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)

  
    

With thehdw.DCVI.Pins("VESDREF_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = False
        .Disconnect
    End With
     thehdw.Wait 1 * ms
  With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms

'**************************************************************************************
'**************************************************************************************


'********************************************************************************************
'********************************************************************************************
'*********************************************************************************************
'*********      continuity on pins M5C2 considering diode vs VSS ( VSS=0 )       **************

With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

    DC_M5C2_vs_VSS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe)

'   thehdw.Wait 1 * ms

   
 With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
'        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply '''f_V_0v_2ma_M_I_2ma
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
    thehdw.Wait 1 * ms
        .Gate = False
        .Disconnect
    End With
     thehdw.Wait 1 * ms
  
     Call TestLimitsinflow(DC_M5C2_vs_VSS)
       Call TestLimitsinflow(DC_M5C2_vs_VSS)
 
'*********************************************************************************************
'*********************************************************************************************
'**************************************************************************************
'**************************************************************************************


'*********************************************************************************************
'*********      continuity on pins VSSD considering diode vs VDDD ( VDDD=0 )       ******************

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms
     
    DC_VSSD_vs_VDDD = thehdw.DCVI.Pins("VSSD_DC30").Meter.Read(tlStrobe)

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms
     
        Call TestLimitsinflow(DC_VSSD_vs_VDDD)
        Call TestLimitsinflow(DC_VSSD_vs_VDDD)


'*********************************************************************************************


'disconnect
    thehdw.DCVI.Pins("DC30_CONT_ALL").Disconnect (tlDCVIConnectDefault)
    thehdw.DCVI.Pins("DC30_CONT_ALL,VSSCP_DC30,VSSD_DC30").Alarm(tlDCVIAlarmMode) = tlAlarmDefault

    thehdw.PPMU.Pins("HSD_CONT_ALL").Disconnect
    
     With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms


     With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
        .Gate = False
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    

    

    SetDatabitsOff ("K50, K6,K16, K17,K4")
'    SetDatabitsOff ("K16, K17")
'    SetDatabitsOff ("K4")
'
'    SetDatabitsOff ("K4") '**** connect VSSD and VSSCP to dc30; it may be considered as normal pins ( testing continuity )
'    thehdw.Wait 1 * ms
'
      With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = False
        .Disconnect
    End With
     thehdw.Wait 1 * ms
    
    thehdw.Digital.DisconnectPins ("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
    thehdw.Wait 1 * ms
   
    Exit Function
errHandler:
        If AbortTest Then Exit Function Else Resume Next


End Function




Public Function New_Continuity_post()

Dim RetPinLevels As String
Dim RetTimeSet As String
Dim RetEdgeSet As String
Dim RetDCCategory As String
Dim RetDCSelector As String
Dim RetACCategory As String
Dim RetACSelector As String
Dim Overlay As String

Call TheExec.DataManager.GetInstanceContext(RetDCCategory, RetDCSelector, RetACCategory, RetACSelector, _
                                            RetTimeSet, RetEdgeSet, RetPinLevels, Overlay)

If RetPinLevels <> TL_C_EMPTYSTR Then thehdw.PinLevels.ApplyPower


If RetTimeSet <> TL_C_EMPTYSTR Then Call thehdw.Digital.Timing.Load

Dim CRES_force100ua As New PinListData

Dim DC_Sink_odd As New PinListData
Dim HSD_Sink_odd As New PinListData
Dim HVD_Sink_odd As New PinListData
Dim HVD_Sink_odd_1 As New PinListData
Dim HVD_Sink_odd_2 As New PinListData
Dim HVD_Sink_odd_3 As New PinListData
    
    
Dim DC_Sink_even As New PinListData
Dim HSD_Sink_even As New PinListData
Dim HVD_Sink_even_1 As New PinListData
Dim HVD_Sink_even_2 As New PinListData
    
Dim DC_odd_vs_VDD As New PinListData
Dim DC_even_vs_VDD As New PinListData
Dim HSD_odd_vs_VDD As New PinListData
Dim HSD_even_vs_VDD As New PinListData

Dim DC_VSSCP_vs_VDDCP As New PinListData
Dim HSD_M5C1_vs_VDDCP As New PinListData
    
Dim HSD_PWM_SDI_vs_VDDQ As New PinListData
Dim HSD_SYNC_SCLK_CS_SDO_vs_VDDQ As New PinListData

Dim DC_SNS1_RCIN_RCTB_TAGB_POUTA_vs_VESDREF As New PinListData
Dim HSD_HFD_vs_VESDREF As New PinListData

Dim DC_TAGA_RCTA_SRTN_POUTB_vs_VESDREF As New PinListData
Dim HVD_TEST1_TEST2_vs_VESDREF As New PinListData

Dim DC_M5C2_vs_VSS  As New PinListData

Dim DC_VSSD_vs_VDDD As New PinListData

Dim SiteNum As Variant

Dim NEW_CONT As Boolean
    
'If RetTimeSet <> TL_C_EMPTYSTR Then Call thehdw.Digital.Timing.Load

' Start continuity_pre

' To connect to the pins directly: K50->CN2A, K17->REF, K6->VDDCP
    
   '******************************************************************************
        
    thehdw.DCVI.Pins("VESDREF_DC30").Alarm(tlDCVIAlarmGuard) = tlAlarmOff

    SetDatabitsOff ("K15,K40")
    thehdw.Wait 1 * ms

    thehdw.Digital.DisconnectPins ("SCLK_HSD,TEST1_HSDHV,TEST2_HSDHV,CS_HSD,SDI_HSD,PWM_HSD,SDO_HSD,SYNC_HSD,VDDQ_HSD,VDDCP_HSD")
    thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("SNS1_DC30,HFD_DC30,SRTN_DC30,RCIN_DC30_1,RCTB_DC30,VLA_DC30,VLB_DC30") ' 0V on sns1
        .Voltage = 0
       .Gate = False
       .Disconnect
    End With
    thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VDD_DC30,VDDD_DC30,VDDQ_DC30,VESDREF_DC30")
        .Voltage = 0
       .Gate = False
       .Disconnect
    End With
    thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("VDDCP_DC90A_H,VDDCP_DC90A_L")
        .Voltage = 0
       .Gate = False
       .Disconnect
    End With
   thehdw.Wait 1 * ms
    
    With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 75 * v
        .CurrentRange = 32 * mA
        .Clamp.current.Max = 32 * mA
        .Clamp.current.Min = -32 * mA
        .Voltage = 0 * v
        .Disconnect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    thehdw.Wait 1 * ms
    SetDatabitsOn ("K50, K6,K16,K17,K4")
    thehdw.Wait 1 * ms
    
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 3 * ms

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Connect
    End With
    thehdw.Wait 1 * ms
       
' All HSD pins at 0V, using HVPPMU
    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
        .Voltage = 0
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
     thehdw.Wait 1 * ms
    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
     
    
    SetDatabitsOff ("K1,K3,K5,K8,K60") '''18nov2014,trtz
    
    SetDatabitsOn ("K50, K6,K16, K17,K4")
    thehdw.Wait 1 * ms
        
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 1 * ms

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Connect
    End With
    thehdw.Wait 1 * ms
    
   
' All HSD pins at 0V, using HVPPMU
    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
    
    
If TheExec.CurrentJob = "UK21FA01" Then
' ###########################################################################
'adding a screening test for cres_dc30 to avoid failing at cold test
    Dim CRES_p As New PinListData
    Dim kira As Integer
    Dim tambah As Double
    
    
    thehdw.DCVI.Pins("CRES_DC30").PSets("PLT_p").Apply
    thehdw.Wait 1 * ms

    CRES_p = thehdw.DCVI.Pins("CRES_DC30").Meter.Read(tlStrobe, 10, 100000, tlDCVIMeterReadingFormatAverage)

   Call TestLimitsinflow(CRES_p) ' new test 501
   
    thehdw.DCVI.Pins("CRES_DC30").Voltage = 0
    thehdw.DCVI.Pins("CRES_DC30").CurrentRange = 2 * mA
    thehdw.Wait 1 * ms
    thehdw.DCVI.Pins("CRES_DC30").current = 2 * mA
    thehdw.Wait 1 * ms

    
' ###########################################################################
End If

If (0) Then ' IF OPEN/SHORT ON CRES_DC30 TO PUT AT 1
    With thehdw.DCVI.Pins("CRES_DC30")
        .Disconnect
    End With
    thehdw.Wait 3 * ms
End If


    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_I_100ua_2v_M_V_2v").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 1 * ms
     
    CRES_force100ua.AddPin ("CRES_DC30")
    thehdw.DCVI.Pins("CRES_DC30").ClearCaptureMemory

    For Each SiteNum In TheExec.Sites.Active
           If TheExec.Sites.Active Then
             CRES_force100ua = thehdw.DCVI.Pins("CRES_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
         Else
         End If
    Next SiteNum
    
    Call TestLimitsinflow(CRES_force100ua)
    Call TestLimitsinflow(CRES_force100ua)

    With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
    End With

    
  thehdw.DCVI.Pins("SNS1_DC30").Alarm(tlDCVIAlarmAll) = tlAlarmDefault
    
    
' ###########################################################################

 ' #############
 '   ODD  PINS
 ' #############

    ' Sink 100ua
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_I_N100ua_10v_M_V_10v").Apply
     thehdw.Wait 1 * ms
    
'''18NOV2014, TRTZ ####################################
    With thehdw.PPMU.Pins("PFGB_HSD")
        .Gate = tlOn
    End With
'''####################################################3
    
    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceI -0.0001
 thehdw.Wait 1 * ms
    
    
   With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = -100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
   End With
     thehdw.Wait 2 * ms
     
   thehdw.DCVI.Pins("DC30_CONT_ODD").ClearCaptureMemory
   thehdw.Wait 1 * ms
   
''    DC_Sink_odd = TheHdw.DCVI.Pins("DC30_CONT_ODD").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
       
    Dim TAGA_DC30_OS As New PinListData
    Dim RCTA_DC30_OS As New PinListData
    Dim SRTN_DC30_OS As New PinListData
    Dim VDDD_DC30_OS As New PinListData
    Dim CN1B_DC30_OS As New PinListData
    Dim CN1A_DC30_1_OS As New PinListData
    Dim VLB_DC30_OS As New PinListData
    Dim VDDCP_DC30_OS As New PinListData
    Dim M5C2_DC30_OS As New PinListData
    Dim POUTB_DC30_1_OS As New PinListData
    Dim CN3_DC30_OS As New PinListData
       
    TAGA_DC30_OS = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    RCTA_DC30_OS = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    SRTN_DC30_OS = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDD_DC30_OS = thehdw.DCVI.Pins("VDDD_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1B_DC30_OS = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN1A_DC30_1_OS = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)

     VLB_DC30_OS = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    VDDCP_DC30_OS = thehdw.DCVI.Pins("VDDCP_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    M5C2_DC30_OS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    POUTB_DC30_1_OS = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)
    CN3_DC30_OS = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe, 200, 100000, tlDCVIMeterReadingFormatAverage)

    Dim PFGB_HSD_OS As New PinListData
    Dim SYNC_HSD_OS As New PinListData
    Dim SDO_HSD_OS As New PinListData
    Dim SDI_HSD_OS As New PinListData
    Dim VDDQ_HSD_OS As New PinListData
    Dim VOUT_HSD_OS As New PinListData
    Dim NCpin35_OS As New PinListData
    Dim NCpin45_OS As New PinListData

    PFGB_HSD_OS = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
    SYNC_HSD_OS = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)
    VDDQ_HSD_OS = thehdw.PPMU.Pins("VDDQ_HSD").Read(tlPPMUReadMeasurements)
    VOUT_HSD_OS = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)
    NCpin35_OS = thehdw.PPMU.Pins("NCpin35").Read(tlPPMUReadMeasurements)
    NCpin45_OS = thehdw.PPMU.Pins("NCpin45").Read(tlPPMUReadMeasurements)


    With thehdw.HVPMU.Pins("HVD_CONT_ODD")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = -100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
   End With
     thehdw.Wait 1 * ms
   
    HVD_Sink_odd_1 = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_odd_2 = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_odd_3 = thehdw.HVPMU.Pins("NCpin23").Read(tlHVPMUVoltage, tlStrobe)
    
    ' Reset
    thehdw.DCVI.Pins("DC30_CONT_ODD").PSets("f_V_0v_2ma_M_I_2ma").Apply
    
    Call TestLimitsinflow(TAGA_DC30_OS)
    Call TestLimitsinflow(TAGA_DC30_OS)
    
    Call TestLimitsinflow(RCTA_DC30_OS)
    Call TestLimitsinflow(RCTA_DC30_OS)
    
    Call TestLimitsinflow(SRTN_DC30_OS)
    Call TestLimitsinflow(SRTN_DC30_OS)
    
    Call TestLimitsinflow(VDDD_DC30_OS)
    Call TestLimitsinflow(VDDD_DC30_OS)
    
    Call TestLimitsinflow(CN1B_DC30_OS)
    Call TestLimitsinflow(CN1B_DC30_OS)
    
    Call TestLimitsinflow(CN1A_DC30_1_OS)
    Call TestLimitsinflow(CN1A_DC30_1_OS)
    
    Call TestLimitsinflow(VLB_DC30_OS)
    Call TestLimitsinflow(VLB_DC30_OS)
    
    Call TestLimitsinflow(VDDCP_DC30_OS)
    Call TestLimitsinflow(VDDCP_DC30_OS)
    
    Call TestLimitsinflow(M5C2_DC30_OS)
    Call TestLimitsinflow(M5C2_DC30_OS)
    
    Call TestLimitsinflow(POUTB_DC30_1_OS)
    Call TestLimitsinflow(POUTB_DC30_1_OS)
    
    Call TestLimitsinflow(CN3_DC30_OS)
    Call TestLimitsinflow(CN3_DC30_OS)
    
    
    Call TestLimitsinflow(PFGB_HSD_OS)
    Call TestLimitsinflow(PFGB_HSD_OS)
    Call TestLimitsinflow(SYNC_HSD_OS)
    Call TestLimitsinflow(SYNC_HSD_OS)
    Call TestLimitsinflow(SDO_HSD_OS)
    Call TestLimitsinflow(SDO_HSD_OS)
    Call TestLimitsinflow(SDI_HSD_OS)
    Call TestLimitsinflow(SDI_HSD_OS)
    Call TestLimitsinflow(VDDQ_HSD_OS)
    Call TestLimitsinflow(VDDQ_HSD_OS)
    Call TestLimitsinflow(VOUT_HSD_OS)
    Call TestLimitsinflow(VOUT_HSD_OS)
    Call TestLimitsinflow(NCpin35_OS)
    Call TestLimitsinflow(NCpin35_OS)
    Call TestLimitsinflow(NCpin45_OS)
    Call TestLimitsinflow(NCpin45_OS)
    
    Call TestLimitsinflow(HVD_Sink_odd_1)
    Call TestLimitsinflow(HVD_Sink_odd_1)
    
    Call TestLimitsinflow(HVD_Sink_odd_2)
    Call TestLimitsinflow(HVD_Sink_odd_2)
    
    Call TestLimitsinflow(HVD_Sink_odd_3)
    Call TestLimitsinflow(HVD_Sink_odd_3)
  
   
    With thehdw.HVPMU.Pins("HVD_CONT_ODD")
'''        .current = 2 * mA
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    
    thehdw.PPMU.Pins("HSD_CONT_ODD").ForceV 0
 thehdw.Wait 1 * ms

' #############
 '   EVEN  PINS
 ' #############

    ' Sink 100ua
    
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_I_N100ua_10v_M_V_10v").Apply
     thehdw.Wait 1 * ms
   
    Dim SNS1_DC30_OS As New PinListData
    Dim RCIN_DC30_1_OS As New PinListData
    Dim RCTB_DC30_OS As New PinListData
    Dim CN2B_DC30_OS As New PinListData
    Dim CN2A_DC30_OS As New PinListData
    Dim REF_DC30_OS As New PinListData
    Dim VLA_DC30_OS As New PinListData
    Dim VESDREF_DC30_OS As New PinListData
    Dim VDD_DC30_OS As New PinListData
    Dim TAGB_DC30_1_OS As New PinListData
    Dim POUTA_DC30_1_OS As New PinListData


    thehdw.DCVI.Pins("VDD_DC30,VESDREF_DC30,REF_DC30,VLA_DC30,CN2A_DC30,CN2B_DC30").ClearCaptureMemory
    thehdw.Wait 1 * ms
   
    SNS1_DC30_OS = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe)
    RCIN_DC30_1_OS = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe)
    RCTB_DC30_OS = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe)
    CN2B_DC30_OS = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
    CN2A_DC30_OS = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
    REF_DC30_OS = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)
    VLA_DC30_OS = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
    VESDREF_DC30_OS = thehdw.DCVI.Pins("VESDREF_DC30").Meter.Read(tlStrobe)
    VDD_DC30_OS = thehdw.DCVI.Pins("VDD_DC30").Meter.Read(tlStrobe)
    TAGB_DC30_1_OS = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe)
    POUTA_DC30_1_OS = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe)
    
    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceI -0.0001
    thehdw.Wait 1 * ms

    Dim PFGA_HSD_OS As New PinListData
    Dim HFD_HSD_OS As New PinListData
    Dim HFG_HSD_OS As New PinListData
    Dim PWM_HSD_OS As New PinListData
    Dim SCLK_HSD_OS As New PinListData
    Dim CS_HSD_OS As New PinListData
    Dim CVF_HSD_OS As New PinListData
    Dim NCpin36_OS As New PinListData
    Dim M5C1_HSD_OS As New PinListData

    PFGA_HSD_OS = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
    HFD_HSD_OS = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements)
    HFG_HSD_OS = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
    PWM_HSD_OS = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
    SCLK_HSD_OS = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
    CS_HSD_OS = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
    CVF_HSD_OS = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)
    NCpin36_OS = thehdw.PPMU.Pins("NCpin36").Read(tlPPMUReadMeasurements)
    M5C1_HSD_OS = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)

'''   HSD_Sink_even = TheHdw.PPMU.Pins("HSD_CONT_EVEN").Read(tlPPMUReadMeasurements)
    
    With thehdw.HVPMU.Pins("NCpin22,NCpin24")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = -100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 3 * ms

    HVD_Sink_even_1 = thehdw.HVPMU.Pins("NCpin22").Read(tlHVPMUVoltage, tlStrobe)
    HVD_Sink_even_2 = thehdw.HVPMU.Pins("NCpin24").Read(tlHVPMUVoltage, tlStrobe)
    
' Reset
    thehdw.DCVI.Pins("DC30_CONT_EVEN").PSets("f_V_0v_2ma_M_I_2ma").Apply
'    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .Voltage = 0
        .Disconnect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With

    thehdw.PPMU.Pins("HSD_CONT_EVEN").ForceV 0

    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(SNS1_DC30_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCIN_DC30_1_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(RCTB_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2B_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(CN2A_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(REF_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VLA_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VESDREF_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(VDD_DC30_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(TAGB_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)
    Call TestLimitsinflow(POUTA_DC30_1_OS)
    
    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(PFGA_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFD_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(HFG_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(PWM_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(SCLK_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CS_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(CVF_HSD_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(NCpin36_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    Call TestLimitsinflow(M5C1_HSD_OS)
    
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_1)
    Call TestLimitsinflow(HVD_Sink_even_2)
    Call TestLimitsinflow(HVD_Sink_even_2)

' **************************************************************************************

' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Disconnect
    End With
    thehdw.Wait 1 * ms
    
    With thehdw.HVPMU.Pins("HVD_CONT_ALL")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms
    
'*********************************************************************************************
'*********      continuity on pins considering diode vs VDD ( VDD=0 )       ******************

'**** even

    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("DCcontEVEN_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("HSDcontEVEN_vs_VDD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms

    Dim VLA_DC30_OS_VDD As New PinListData
    Dim CN2B_DC30_OS_VDD As New PinListData
    Dim CN2A_DC30_OS_VDD As New PinListData
    Dim REF_DC30_OS_VDD As New PinListData
    
    VLA_DC30_OS_VDD = thehdw.DCVI.Pins("VLA_DC30").Meter.Read(tlStrobe)
    CN2B_DC30_OS_VDD = thehdw.DCVI.Pins("CN2B_DC30").Meter.Read(tlStrobe)
    CN2A_DC30_OS_VDD = thehdw.DCVI.Pins("CN2A_DC30").Meter.Read(tlStrobe)
    REF_DC30_OS_VDD = thehdw.DCVI.Pins("REF_DC30").Meter.Read(tlStrobe)
    
    '''   DC_even_vs_VDD = TheHdw.DCVI.Pins("DCcontEVEN_vs_VDD").Meter.Read(tlStrobe)
    
    Dim PFGA_HSD_OS_VDD As New PinListData
    Dim HFG_HSD_OS_VDD As New PinListData
    Dim CVF_HSD_OS_VDD As New PinListData
    
    PFGA_HSD_OS_VDD = thehdw.PPMU.Pins("PFGA_HSD").Read(tlPPMUReadMeasurements)
    HFG_HSD_OS_VDD = thehdw.PPMU.Pins("HFG_HSD").Read(tlPPMUReadMeasurements)
    CVF_HSD_OS_VDD = thehdw.PPMU.Pins("CVF_HSD").Read(tlPPMUReadMeasurements)
    
    '''   HSD_even_vs_VDD = TheHdw.PPMU.Pins("HSDcontEVEN_vs_VDD").Read(tlPPMUReadMeasurements)

    Call TestLimitsinflow(VLA_DC30_OS_VDD)
    Call TestLimitsinflow(VLA_DC30_OS_VDD)
    Call TestLimitsinflow(CN2B_DC30_OS_VDD)
    Call TestLimitsinflow(CN2B_DC30_OS_VDD)
    Call TestLimitsinflow(CN2A_DC30_OS_VDD)
    Call TestLimitsinflow(CN2A_DC30_OS_VDD)
    Call TestLimitsinflow(REF_DC30_OS_VDD)
    Call TestLimitsinflow(REF_DC30_OS_VDD)
    
    Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(PFGA_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
    Call TestLimitsinflow(HFG_HSD_OS_VDD)
    Call TestLimitsinflow(CVF_HSD_OS_VDD)
    Call TestLimitsinflow(CVF_HSD_OS_VDD)

   '**** reset

' **************************************************************************************
    
' All DC Pins set to 0v
    With thehdw.DCVI.Pins("DC30_CONT_ALL,VSSD_DC30,VSSCP_DC30") '**** without CRES
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms

' All HSD pins at 0V, using PPMU
    With thehdw.PPMU.Pins("HSD_CONT_ALL")
        .ForceV 0
        .Disconnect
    End With
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    thehdw.Wait 1 * ms

'********************************************************************************
    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("DCcontODD_vs_VDD") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms
    
'***** odd

    Dim CN1B_DC30_OS_VDD As New PinListData
    Dim CN3_DC30_OS_VDD As New PinListData
    Dim CN1A_DC30_1_OS_VDD As New PinListData
    Dim VLB_DC30_OS_VDD As New PinListData

    CN1B_DC30_OS_VDD = thehdw.DCVI.Pins("CN1B_DC30").Meter.Read(tlStrobe)
    CN3_DC30_OS_VDD = thehdw.DCVI.Pins("CN3_DC30").Meter.Read(tlStrobe)
    CN1A_DC30_1_OS_VDD = thehdw.DCVI.Pins("CN1A_DC30_1").Meter.Read(tlStrobe)
    VLB_DC30_OS_VDD = thehdw.DCVI.Pins("VLB_DC30").Meter.Read(tlStrobe)

    Dim PFGB_HSD_OS_VDD As New PinListData
    Dim VOUT_HSD_OS_VDD As New PinListData

    PFGB_HSD_OS_VDD = thehdw.PPMU.Pins("PFGB_HSD").Read(tlPPMUReadMeasurements)
    VOUT_HSD_OS_VDD = thehdw.PPMU.Pins("VOUT_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.DCVI.Pins("DCcontODD_vs_VDD")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Disconnect
        .Gate = False
    End With

    With thehdw.PPMU.Pins("HSDcontODD_vs_VDD")
        .ForceI 0.0001
        .Disconnect
    End With
    
    With thehdw.DCVI.Pins("VDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    
    Call TestLimitsinflow(CN1B_DC30_OS_VDD)
    Call TestLimitsinflow(CN1B_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN3_DC30_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
    Call TestLimitsinflow(CN1A_DC30_1_OS_VDD)
    Call TestLimitsinflow(VLB_DC30_OS_VDD)
    Call TestLimitsinflow(VLB_DC30_OS_VDD)
    
    Call TestLimitsinflow(PFGB_HSD_OS_VDD)
    Call TestLimitsinflow(PFGB_HSD_OS_VDD)
    Call TestLimitsinflow(VOUT_HSD_OS_VDD)
    Call TestLimitsinflow(VOUT_HSD_OS_VDD)

'********************************************************************************************
'********************************************************************************************


'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDCP ( VDDCP=0 )       ******************

    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("M5C1_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms

    DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe, 100, 100000)

' Instable on site0 during GM debug also with OLD rev; HW issue???
'    For Each SiteNum In TheExec.Sites.Active
'           If DC_VSSCP_vs_VDDCP < 300 * mV Then
'            thehdw.DCVI.Pins("VSSCP_DC30").ClearCaptureMemory
'
'            DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe)
'         End If
'    Next SiteNum
'        thehdw.Wait 3 * ms
'    For Each SiteNum In TheExec.Sites.Active
'           If DC_VSSCP_vs_VDDCP < 300 * mV Then
'            thehdw.DCVI.Pins("VSSCP_DC30").ClearCaptureMemory
'
'            DC_VSSCP_vs_VDDCP = thehdw.DCVI.Pins("VSSCP_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
'         End If
'    Next SiteNum
' Instable on site0 during GM debug also with OLD rev; HW issue???
    
    HSD_M5C1_vs_VDDCP = thehdw.PPMU.Pins("M5C1_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.DCVI.Pins("VDDCP_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.DCVI.Pins("VSSCP_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
     
    With thehdw.PPMU.Pins("M5C1_HSD")
        .ForceI 0.0001
        .Disconnect
    End With

    Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
    Call TestLimitsinflow(DC_VSSCP_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)
    Call TestLimitsinflow(HSD_M5C1_vs_VDDCP)

'*********************************************************************************************
'*********      continuity on pins considering diode vs VDDQ ( VDDQ=0 )       **************

    With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms
    
'***** even

   Dim PWM_HSD_OS_VDDQ As New PinListData
   Dim SDI_HSD_OS_VDDQ As New PinListData

 
    PWM_HSD_OS_VDDQ = thehdw.PPMU.Pins("PWM_HSD").Read(tlPPMUReadMeasurements)
    SDI_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDI_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.PPMU.Pins("PWM_HSD,SDI_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
    thehdw.Wait 1 * ms
'****** odd

 With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms

 Dim SYNC_HSD_OS_VDDQ As New PinListData
  Dim SCLK_HSD_OS_VDDQ As New PinListData
 Dim CS_HSD_OS_VDDQ As New PinListData
  Dim SDO_HSD_OS_VDDQ As New PinListData

 
    SYNC_HSD_OS_VDDQ = thehdw.PPMU.Pins("SYNC_HSD").Read(tlPPMUReadMeasurements)
    SCLK_HSD_OS_VDDQ = thehdw.PPMU.Pins("SCLK_HSD").Read(tlPPMUReadMeasurements)
    CS_HSD_OS_VDDQ = thehdw.PPMU.Pins("CS_HSD").Read(tlPPMUReadMeasurements)
    SDO_HSD_OS_VDDQ = thehdw.PPMU.Pins("SDO_HSD").Read(tlPPMUReadMeasurements)

'''    HSD_SYNC_SCLK_CS_SDO_vs_VDDQ = TheHdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD").Read(tlPPMUReadMeasurements)

    With thehdw.PPMU.Pins("SYNC_HSD,SCLK_HSD,CS_HSD,SDO_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
'    thehdw.Wait 1 * ms
    
     With thehdw.DCVI.Pins("VDDQ_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms
    
    Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
    Call TestLimitsinflow(PWM_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDI_HSD_OS_VDDQ)
    
    Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
    Call TestLimitsinflow(SYNC_HSD_OS_VDDQ)
    Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
    Call TestLimitsinflow(SCLK_HSD_OS_VDDQ)
    Call TestLimitsinflow(CS_HSD_OS_VDDQ)
    Call TestLimitsinflow(CS_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDO_HSD_OS_VDDQ)
    Call TestLimitsinflow(SDO_HSD_OS_VDDQ)

  
   
'*********************************************************************************************
'*********      continuity on pins considering diode vs VESDREF ( VESDREF=0 )       **************
  With thehdw.HVPMU.Pins("VESDREF_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 75 * v
        .CurrentRange = 32 * mA
        .Clamp.current.Max = 32 * mA
        .Clamp.current.Min = -32 * mA
        .Voltage = 0 * v
        .Connect (tlHVPMUConnectForce + tlHVPMUConnectKelvin)
    End With
    'thehdw.Wait 3 * ms

    With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    'thehdw.Wait 1 * ms

    With thehdw.PPMU.Pins("HFD_HSD")
        .ForceI 0.0001
        .Connect
    End With
    thehdw.Wait 3 * ms
    
'**** even
    
    Dim SNS1_DC30_OS_VESDREF As New PinListData
    Dim RCIN_DC30_1_OS_VESDREF As New PinListData
    Dim RCTB_DC30_OS_VESDREF As New PinListData
    Dim TAGB_DC30_1_OS_VESDREF As New PinListData
    Dim POUTA_DC30_1_OS_VESDREF As New PinListData
    
    SNS1_DC30_OS_VESDREF = thehdw.DCVI.Pins("SNS1_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCIN_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("RCIN_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    RCTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTB_DC30").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    TAGB_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("TAGB_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)
    POUTA_DC30_1_OS_VESDREF = thehdw.DCVI.Pins("POUTA_DC30_1").Meter.Read(tlStrobe, 100, 100000, tlDCVIMeterReadingFormatAverage)

   HSD_HFD_vs_VESDREF = thehdw.PPMU.Pins("HFD_HSD").Read(tlPPMUReadMeasurements, 100)


    With thehdw.DCVI.Pins("SNS1_DC30,RCIN_DC30_1,RCTB_DC30,TAGB_DC30_1,POUTA_DC30_1")
        .PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
        .Gate = False
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
    
    With thehdw.PPMU.Pins("HFD_HSD")
        .ForceI 0.0001
        .Disconnect
    End With
    'thehdw.Wait 1 * ms

  Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
   Call TestLimitsinflow(SNS1_DC30_OS_VESDREF)
  Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(RCIN_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
  Call TestLimitsinflow(RCTB_DC30_OS_VESDREF)
  Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(TAGB_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
  Call TestLimitsinflow(POUTA_DC30_1_OS_VESDREF)
   
   Call TestLimitsinflow(HSD_HFD_vs_VESDREF)
   Call TestLimitsinflow(HSD_HFD_vs_VESDREF)


'**************************************************************************************
'**************************************************************************************

'********************************************************************************************
'********************************************************************************************


    With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 3 * ms
     
     
'**** odd
    thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
    'thehdw.Wait 1 * ms

    
   With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23")
        .Mode = tlHVPMUModeCurrent
        .VoltageRange = 2
        .CurrentRange = 100 * uA
        .current = 100 * uA
        .Clamp.Voltage.Max = 3 * v
        .Clamp.Voltage.Min = -3 * v
        .Connect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
   End With
    thehdw.Wait 3 * ms
    
    Dim TAGA_DC30_OS_VESDREF As New PinListData
    Dim RCTA_DC30_OS_VESDREF As New PinListData
    Dim SRTN_DC30_OS_VESDREF As New PinListData
    Dim POUTB_DC30_OS_VESDREF As New PinListData


    TAGA_DC30_OS_VESDREF = thehdw.DCVI.Pins("TAGA_DC30").Meter.Read(tlStrobe)
    RCTA_DC30_OS_VESDREF = thehdw.DCVI.Pins("RCTA_DC30").Meter.Read(tlStrobe)
    SRTN_DC30_OS_VESDREF = thehdw.DCVI.Pins("SRTN_DC30").Meter.Read(tlStrobe)
    POUTB_DC30_OS_VESDREF = thehdw.DCVI.Pins("POUTB_DC30").Meter.Read(tlStrobe)

    Dim TEST1_HSDHV_OS_VESDREF As New PinListData
    Dim TEST2_HSDHV_OS_VESDREF As New PinListData

    TEST1_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST1_HSDHV").Read(tlHVPMUVoltage, tlStrobe)
    TEST2_HSDHV_OS_VESDREF = thehdw.HVPMU.Pins("TEST2_HSDHV").Read(tlHVPMUVoltage, tlStrobe)

   thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30").PSets("f_V_0v_200ma_M_I_200ma_xCONT").Apply
    thehdw.Wait 1 * ms

    With thehdw.HVPMU.Pins("TEST1_HSDHV,TEST2_HSDHV,NCpin23,NCpin22,NCpin24")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
    'thehdw.Wait 1 * ms

    Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
     Call TestLimitsinflow(TAGA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(RCTA_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(SRTN_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)
    Call TestLimitsinflow(POUTB_DC30_OS_VESDREF)
   
    
    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST1_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)
    Call TestLimitsinflow(TEST2_HSDHV_OS_VESDREF)

    With thehdw.DCVI.Pins("VESDREF_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     'thehdw.Wait 1 * ms
    With thehdw.DCVI.Pins("TAGA_DC30,RCTA_DC30,SRTN_DC30,POUTB_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
     thehdw.Wait 1 * ms

'**************************************************************************************
'**************************************************************************************


'********************************************************************************************
'********************************************************************************************
'*********************************************************************************************
'*********      continuity on pins M5C2 considering diode vs VSS ( VSS=0 )       **************

    With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     'thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
     thehdw.Wait 3 * ms

    DC_M5C2_vs_VSS = thehdw.DCVI.Pins("M5C2_DC30").Meter.Read(tlStrobe)
    
    With thehdw.DCVI.Pins("VSS_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
     'thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("M5C2_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms
  
     Call TestLimitsinflow(DC_M5C2_vs_VSS)
    Call TestLimitsinflow(DC_M5C2_vs_VSS)
 
'*********************************************************************************************
'*********************************************************************************************
'**************************************************************************************
'**************************************************************************************


'*********************************************************************************************
'*********      continuity on pins VSSD considering diode vs VDDD ( VDDD=0 )       ******************

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Connect
    End With
     'thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Connect
    End With
    thehdw.Wait 3 * ms
     
    DC_VSSD_vs_VDDD = thehdw.DCVI.Pins("VSSD_DC30").Meter.Read(tlStrobe)

    With thehdw.DCVI.Pins("VDDD_DC30")
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms

    With thehdw.DCVI.Pins("VSSD_DC30") '**** considering as normal pins
        .PSets("f_I_100ua_2v_M_V_2v_xCONT").Apply
        .Gate = True
        .Disconnect
    End With
'     thehdw.Wait 1 * ms
     
        Call TestLimitsinflow(DC_VSSD_vs_VDDD)
        Call TestLimitsinflow(DC_VSSD_vs_VDDD)


'*********************************************************************************************


'disconnect
    thehdw.DCVI.Pins("DC30_CONT_ALL").Disconnect (tlDCVIConnectDefault)
    
    thehdw.PPMU.Pins("HSD_CONT_ALL").Disconnect
    thehdw.PPMU.Pins("HSD_CONT_ALL").Gate = tlOff ''''18NOV2014
    
     With thehdw.HVPMU.Pins("NCpin22,NCpin23,NCpin24,TEST1_HSDHV,TEST2_HSDHV")
        .Mode = tlHVPMUModeVoltage
        .VoltageRange = 0.5
        .CurrentRange = 100 * uA
        .Voltage = 0
        .Clamp.current.Max = 100 * uA
        .Clamp.current.Min = -100 * uA
        .Disconnect (tlHVPMUConnectKelvin + tlHVPMUConnectForce)
    End With
'    thehdw.Wait 1 * ms
    
    
     With thehdw.DCVI.Pins("CRES_DC30")
        .PSets("f_V_0v_200ma_M_I_200ma").Apply
        .Gate = False
        .Disconnect
    End With
    'thehdw.Wait 1 * ms
    
    SetDatabitsOff ("K50, K6,K16, K17,K4")
    'thehdw.Wait 1 * ms
    
      With thehdw.DCVI.Pins("VSSD_DC30,VSSCP_DC30") '**** considering as normal pins
        .PSets("f_V_0v_2ma_M_I_2ma").Apply
        .Gate = False
        .Disconnect
    End With
    thehdw.Wait 3 * ms
     

     

End Function






