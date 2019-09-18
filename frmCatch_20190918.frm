VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCatch 
   Caption         =   " Phosphorus Land Use and Slope - PLUS+"
   ClientHeight    =   11220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   OleObjectBlob   =   "frmCatch_20190918.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'NOTE CHECK THE SETTING FOR M10 ETC. BEFORE DOING ANYTHING ELSE

Option Explicit
'RunBatchProcesses
'#################################################################################################################################
'#################################################################################################################################
'########## Written by:                                       ####################################################################
'########## David Donnelly                                    ####################################################################
'########## The James Hutton Institute                        ####################################################################
'########## Craigiebuckler, Aberdeen, AB15 8QH                ####################################################################
'########## david.donnelly@hutton.ac.uk                       ####################################################################
'########## 01224 395265                                      ####################################################################
'#################################################################################################################################
'########## Application to model Phosphorus in                ####################################################################
'########## networked catchments, produced for                ####################################################################
'########## Scottish Environment Protection                   ####################################################################
'########## Agency (SEPA) & JHI                               ####################################################################
'##########                                                   ####################################################################
'
'Copyright (C) 2012-15  David Donnelly, The James Hutton Institute

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, version 3 of the License.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'########## While all appropriate checks and precautions have  ###################################################################
'########## been undertaken during design, execution and      ####################################################################
'########## testing of PLUS+, The James Hutton Institute      ####################################################################
'########## cannot be held responsible for any consequences   ####################################################################
'########## arising from the use or mis-use of the model, or  ####################################################################
'########## any part therein. Users of PLUS+ must ensure that ####################################################################
'########## results obtained from PLUS+ are cross-checked.    ####################################################################
'#################################################################################################################################
'########## This code incorporates public domain scripts and  ####################################################################
'########## ESRI sample code.                                 ####################################################################
'#################################################################################################################################
'########## Release number: 1.0 Release date  : 08/12/2009    ####################################################################
'########## Release number: 1.1 Release date  : 12/07/2010    ####################################################################
'##########  Revision List: added loading of per capita TP loads from scenario                                             #######
'##########                 modifications to take account of new land cover in a catchment, including scenario saving      #######
'##########                 user polygons now access underlying slope data to obtain export coefficients                   #######
'##########                 P load of stand alone catchments calculated and displayed                                      #######
'##########                 error fixed in checking on all tables being loaded                                             #######
'##########                 addition of reporting function                                                                 #######
'##########                 addition of upstream inputs to the output window                                               #######
'##########                 changed the WFD text to “high, good, moderate, poor, bad”                                      #######
'##########                 incorporated SEPA_Loch_WB_classification, including GBLakes to WBID lookup                     #######
'##########                 incorporated ability to select WBID or GBLakes ID with menus automatically syncing             #######
'##########                 - the tool still uses the GB Lakes ID for processing as that is what is in the databases       #######
'##########                 - additionally, the look up table only contains 546 records - there are >8000 in the database  #######
'##########                 - so we cannot use WBID as we don't have all the IDs for the GBLake codes                      #######
'##########                 reporting - JPEG, PDF quality and output resolution now working correctly                      #######
'##########                 reporting - now exports to comma separated text files                                          #######
'##########                 implemented additional point sources, requiring new scenario table                             #######
'##########                 fixed error in rural sewage source calculation                                                 #######
'##########                 fixed error in land cover coefficient modifications whereby the coefficients of all the land   #######
'##########                       covers in the chosen catchment were replacing those in connected networks                #######
'########## Release number: 1.2 Release date  : 20/09/2010    ####################################################################
'##########                 changed WB_ID to WFD_WB_ID and GB_WBID to GB_WB_ID as these fields have been renamed           #######
'##########                 CheckReferences routine commented out - it doesn't work on Jonathan's - is it MS Visual Studio?
'##########                 investigated the A4_Portrait.mxt dependency - It's a reference in the MXD supplied by SEPA
'##########                      - tried without the MXT and it doesn't make a difference to the running, so it is now removed
'##########                        from the references
'########## Release number: 1.3 Release date  : 14/07/2011    ####################################################################
'##########                 Minor revisions, released to coincide with presentation of tool at SEPA and to incorporate     #######
'##########                 revised contact details etc.                                                                   #######
'########## Release number: 1.4 Release date  : 06/12/11      ####################################################################
'##########                 Revisions following meeting at SEPA, Perth on 18/07/11                                         #######
'##########                       - replace "sitecode" with GBLakes_ID throughout                                          #######
'##########                       - give an option for the map export to be in landscape or portrait                       #######
'##########                       - extensive rewriting to allow large numbers of P point sources                          #######
'##########                       - additions to output and display to include capacity for upgrade/downgrade              #######
'##########                       - modifications to reporting functions to include capacity for upgrade/downgrade         #######
'########## Release number: 1.5 Release date  : TBC           ####################################################################
'##########                 Minor corrections related to running with modifed versions of the input data and improved      #######
'##########                       working with scenarios                                                                   #######
'##########                       - also - bug fixed so that scenarios can be loaded without having to load base data first#######
'##########                       - also - data checks showed that approx. 100 water bodies had incomplete area and run-off#######
'##########                       - also - these have now been updated in PLUS2 - definitive - so update the data sets too.#######
'########## Release number: 1.6 Release date  : 02/05/2014    ####################################################################
'##########                 Improved documentation and batch processing - this is suitable for use with ArcGIS 10.1        #######
'########## Release number: 1.7 Release date  : after 06/08/2015    ##############################################################
'##########                 further improvements to batch processing -                                                     #######
'##########                 also the addition of the option to change the urban and rural loads globally.                  #######
'##########                 also the inclusion of a table of point sources for non-scenario analysis.                      #######
'########## Release number: 1.8 Release date  : after 29/08/2017    ##############################################################
'##########                 The app was crashing on the GraphicsContainer. Reset action with a new MXD from SEPA
'##########                 This new MXD had a different layout and it appears that the code couldn't find elements
'##########                 within the graphics container at the hard coded location of 0.6007, 0.672 so I have changed
'##########                 those to xmin and ymin and increased the search tolerance to 0.01
'#################################################################################################################################

'This application requires:
'Microsoft Windows Common Controls 6.0 (SP6) - in the library mscomctl.ocx, and it needs to be referenced in Tools->References.
'This enables the following:
            'Microsoft ImageList Control 6.0 (SP4)  (used for the coloured icons to indicate status)
            'Microsoft ListView Control 6.0 (SP4)
'Note that an automated system update in January 2016 over-wrote mscomctl.ocx and causing an automation error stopped PLUS+ from running
'or even from my editing the script.
'Reverting to an older version of this file (15.04.2014) solved the problem.
'The folder c:\temp is used to write temporary files, please ensure it exists

'Note: following certain MS updates an error may be created on loading ArcMap as follows:
'Microsoft Visual Basic
'Object library invalid or contains references to object definitions that could not be found

'The solution to this is to go to C:\Documents and Settings\"UserName"\Application Data\Microsoft\Forms
'and delete the contents, these will be recreated as needed.

'If an error "Compile error: Automation error" occurs and prevents the running of any part of PLUS+
'this may be as a result of mscomctl.ocx becoming corrupted. Remove the frmCatch,
'and reinstall the library using the executable in the Word file.

'#################################################################################
'Note: All field names in this code are hardcoded, but have error traps to highlight problems for user rectification.
'CatchNetRship(i, 0) = GBLAKES_ID
'CatchNetRship(i, 1) = relationship to chosen
'CatchNetRship(i, 2) = jA(L), local runoff in m/yr x local catchment area in m2
'CatchNetRship(i, 3) = sum of the current and upstream local runoffs
'CatchNetRship(i, 4) = marker to record that GBLAKES_ID has been processed in the local runoff calculation (1=true)
'CatchNetRship(i, 5) = Loch volume (m3)
'CatchNetRship(i, 6) = Tw water residence time - =area of loch * mean depth in metres / runoff (i.e. runoff = CatchNetRship(i, 3))
'CatchNetRship(i, 7) = OECDDenominator
'CatchNetRship(i, 8) = OECDExponentDenominator
'CatchNetRship(i, 9) = Sum P in tblCatchP for all of the landcover types for each GBLAKES_ID
'CatchNetRship(i, 10 = Sum of the current and immediate upstream P (total P in the loch in kg)
'CatchNetRship(i, 11)= marker to record that GBLAKES_ID has been processed in the sum P calculation (1=true)
'CatchNetRship(i, 12)= the catchment that this one flows into
'CatchNetRship(i, 13)= the order of this catchment
'CatchNetRship(i, 14)= TP (total loch phosphorus concentration in ug/l)
'CatchNetRship(i, 15)= J (up)
'CatchNetRship(i, 16)= Urban Load
'CatchNetRship(i, 17)= Rural Load
'CatchNetRship(i, 18)= Urban Pop from CatchmentSewage table
'CatchNetRship(i, 19)= Rural Pop from CatchmentSewage table
'CatchNetRship(i, 20)= TP breakpoints reference type
'CatchNetRship(i, 21)= Boundary status
'CatchNetRship(i, 22)= Amount of user entered point source
'CatchNetRship(i, 23)= Loch mean depth
'CatchNetRship(i, 24)= Area of loch
'CatchNetRship(i, 25)= SEPA measured concentration
'CatchNetRship(i, 26)= Back calculated J (total P load) for the SEPA concentration

'The SEPA supplied map template SEPA_A4.mxd is used to generate the output maps. If the template is modified the output maps are likely to be adversely affected
'data required from input tables to populate scenarios
'tblCatchP - includes the LCOVDESC, P and Area for each landcover for each catchment - fundamental for scenario building. Has been built from the exports - varying slopes!

'shapefiles and tables used and relevant attributes:
'1. LocalCatchment_and_Network (a shapefile): GBLAKES_ID, Sitename, Order_, Catch_Net (this number identifies which catchments are connected - 0 means no connections)
'2. LoadPrecursor (a table): GBLAKES_ID, LochOrder, LochArea, LochMeanDepth, LocalArea, OECDDenominator, OECDExponentDenominator, LocalRunoff
'3. CatchmentSewage (a table): GBLAKES_ID, Urb_Rur, Load, Pop,
'4. CatchP (a table): GBLAKES_ID, LCOVDESC, P, Area
'5. TPBreakPoints (a table): GBLAKES_ID, Reference_Type, HighGood_P, GoodModerate_P, ModeratePoor_P, PoorBad_P
'6. Exports (a table): LCOVCODE, SlopeCode, MatchCode, LCOVDESC, Min, Max, Average
'7. PerCapitaTPLoads (a table): Urb_Rur, PerCapitaTPLoad
'8. SlopeClass_LandCover (a shapefile) - to determine the SlopeClass_LandCover for user added polygons

'for scenario creation the following may be modified:
'1. LocalCatchment_and_Network (a shapefile): none
'2. LoadPrecursor (a table): possibly: OECDDenominator, OECDExponentDenominator, LocalRunoff
'3. CatchmentSewage (a table): Urb_Rur, Load, Pop *this is dependent on PerCapitaTPLoads
'4. CatchP (a table): LCOVDESC, P, Area *this is the main change
'5. TPBreakPoints (a table): none
'6. Exports (a table): none - this is a table of coefficients - the user can create a new one if required.
'       Note this would have to be processed into a new SlopeClass_LandCover. This table is used to populate various areas when new land cover is implemented
'7. PerCapitaTPLoads (a table): Urb_Rur, PerCapitaTPLoad *changes here must be cascaded into CatchmentSewage
'8. SlopeClass_LandCover (a shapefile): none

'SEPA loch classification table (user selectable)
'The following field names are used: WATER_BODY_ID, CLASSIFICATION_YEAR, STATUS
'The sequence of fields in the table is not relevant

'To translate between GBLakes ID and WBID the SEPA table GBLakes_WBID_lookup table is used (user selectable)
'The following field names are used: WFD_WB_ID, GB_WB_ID
'The sequence of fields in the table is not relevant


'############################Batch Processing'############################
'When running a batch process of the whole country enable RunBatchProcesses at the end of the Activate module
' - this initiates cmdCreateReportBatch   This process has the image exporting disabled, so no maps will be produced
'You will also probably need to modify the output file lines to a path suitable for your PC:
'        txtOutputReport.Text = "c:\temp\P\" & pRow.Value(pTableCatchment.FindField("SITECODE"))
'        txtOutputFile.Text = "c:\temp\P\" & pRow.Value(pTableCatchment.FindField("SITECODE"))


'This version of the code implements an option to use a modified section of the input data - this uses non-standard land cover
'By default this is disabled, so you need change nothing and it will work normally
'
'Do a search for 'Skene' to get to the modified section
'

Dim pMxDoc As IMxDocument
Dim pMap As IMap
Dim pMxApp As IMxApplication2
Dim pTabColl As IStandaloneTableCollection
Dim fso As Object
Dim pFlowRoutingTable As IStandaloneTable
Dim pSEPAmonitoringTable As IStandaloneTable
Dim pSEPAmonitoringArray() As Variant
Dim pSEPAClassConcStatTable As IStandaloneTable
Dim pSEPAClassConcStatArray() As Variant
Dim dblJ_for_meas_TP As Double
Dim dblSEPA_meas_conc As Double
Dim pGBLakes_WBID_Table As IStandaloneTable
Dim pGBLakes_WBID_Array() As Variant
Dim pFClassCatchment As IFeatureClass
Dim pFLayerCatchment As IFeatureLayer
Dim pDisplayTableCatchment As IDisplayTable
Dim pTableCatchment As ITable
Dim pFieldsCatchment As IFields2
Dim pFieldCatchmentGBLAKES_ID As IField
Dim pFieldCatchmentSiteName As IField
Dim pFieldCatchmentNetwork As IField
Dim lonGBLAKES_IDArray() As Long
Dim lonGBLAKES_IDWithMaxOrder As Long
Dim strSitenameArray() As String
Dim lonGBLAKES_IDNetworkMatchArray() As Long
Dim lonOrderMatchArray() As Long
Dim arrayFlowRouting() As Variant
Dim CatchNetRship() As Variant
Dim CatchmentNetworkFlowsToOnly() As Variant
Dim varSelectedCatchmentCatchP() As Variant
Dim varCatchmentSewage() As Variant 'this contains the read-in data
Dim varArrayExports(89, 1) As Variant 'this contains the read-in data
Dim varPointSource() As Variant 'this contains the read-in data
Dim intOrderArray() As Integer
Dim intMatchingGBLAKES_IDs As Integer
Dim intIndexSelectedGBLAKES_ID As Long
Dim lonNetworkArray() As Long
Dim lonNumGBLAKES_IDs As Long
Dim lonFlowRoutes As Long
Dim intFieldNum As Integer
Dim strChosenSitename As String
Dim lonChosenGBLAKES_ID As Long
Dim lonChosenNetwork As Long
Dim dblCatchP() As Variant
Dim lontblCatchPRecords As Long
Dim intCatchP_LCOVDESCField As Integer
Dim intCatchP_PField As Integer
Dim intCatchP_AreaField As Integer
Dim intCatchP_GBLAKES_IDField As Integer
Dim intMatchCode As Integer
Dim intLCOVCODE As Integer
Dim intSlopeCode As Integer
Dim intLCOVDESC As Integer
Dim intMin As Integer
Dim intMax As Integer
Dim intAverage As Integer
Dim intPerCapitaUrb_RurField As Integer
Dim intPerCapitaTPLoadField As Integer
Dim dblUrbanPerCapitaTPLoad As Double
Dim dblRuralPerCapitaTPLoad As Double
Dim intLowerDensityField As Integer
Dim intUpperDensityField As Integer
Dim strDiscrepancyInClasses As String

Dim intGBLAKES_IDFieldBreakPoints As Integer
Dim intReference_TypeField As Integer
Dim intHighGood_PField As Integer
Dim intGoodModerate_PField As Integer
Dim intModeratePoor_PField As Integer
Dim intPoorBad_PField As Integer

Dim dblTP As Double
Dim dblTP_Up As Double
Dim dblJ_Up As Double
Dim dblJSelectedCatchment As Double
Dim dblLocal_and_Upstream_Runoff As Double
Dim maxOrder As Integer
Dim blnSelectedIsOrderZero As Boolean
Dim dblDeep_a As Double
Dim dblDeep_b As Double
Dim dblShallow_a As Double
Dim dblShallow_b As Double
Dim dblShallowDeep As Double
Dim intLocalRunoffField As Integer
Dim intLocalAreaField As Integer
Dim intPrecursorGBLAKES_IDField As Integer
Dim intLochDepthField As Integer
Dim intLochAreaField As Integer
Dim intOECDDenominatorField As Integer
Dim intOECDExponentDenominatorField As Integer
Dim intOrderField As Integer
Dim intUrb_RurField As Integer
Dim intLoadField As Integer
Dim intPopulationField As Integer
Dim intSewageGBLAKES_IDField As Integer
Dim intPS_GBLAKES_ID_field As Integer
Dim intPS_Type_field As Integer
Dim intPS_Amount_field As Integer

Dim strTblLoadPrecursorName As String
Dim strTblCatchPName As String
Dim strTblExportsName As String
Dim strTblCatchmentSewageName As String
Dim strTblPerCapitaTPLoadsName As String
Dim strTblFlowRouting As String
Dim strTblTPBreakPointsName As String
Dim strTblPointSource As String
Dim strShapefileSlopeClass_LandCover As String
Dim strScenarioLocalCatchmentAndNetwork As String
Dim strScenarioTblCatchmentSewageName As String
Dim strScenarioTblCatchPName As String
Dim strScenarioTblExportsName As String
Dim strScenarioTblLoadPrecursorName As String
Dim strScenarioTblPerCapitaTPLoadsName As String
Dim strScenarioTblTPBreakPointsName As String
Dim strScenarioTbl As String
Dim strScenarioTableNames(6) As String
Dim lonSelectedScenario As Long
Dim pFLayerCatchment_Scenario As IFeatureLayer2
Dim pFClassCatchment_Scenario As IFeatureClass
Dim pScenarioLocalCatchmentAndNetworkTable As ITable
Dim lonScenarioGBLAKES_IDs As Long
Dim blnDataLoadedFromAScenario As Boolean
Dim strLcoverForNetworkChange As String

Dim pCatchmentSewageTable As IStandaloneTable
Dim pCatchPTable As IStandaloneTable
Dim pPointSourceTable As IStandaloneTable 'added on 11.08.2015 in anticipation of WWTW data
Dim blnLoadNonScenarioPointSources As Boolean
Dim arrayCatchPforChosenGBLAKES_ID() As Variant
Dim arrayCatchPforChosenGBLAKES_ID_With_Summary() As Variant
Dim pExportsTable As IStandaloneTable
Dim varArrayExportsTable() As Variant
Dim pLoadPrecursorTable As IStandaloneTable
Dim pPerCapitaTPLoads As IStandaloneTable
Dim arrayPerCapitaTPLoads() As Variant
Dim pTPBreakPoints As IStandaloneTable
Dim blnModifySewageLoad As Boolean
Dim blnModifyOtherPointSourceLoad As Boolean
Dim intScenarioIDField As Integer
Dim intScenarioNameField As Integer
Dim intScenarioCreatorField As Integer
Dim intScenarioCreationDateField As Integer
Dim intScenarioCommentField As Integer
Dim intScenarioRegionField As Integer
Dim strListofGDBContainingScenarioTables() As String
Dim blnScenarioCanBeSaved As Boolean
Dim blnScenarioLoaded As Boolean
Dim strFieldGRIDCODE As String
Dim strFieldAverageExport As String
Dim blnNoLongerBaseline As Boolean

Dim pCatchmentSewageTable_S As IStandaloneTable
Dim pCatchPTable_S As IStandaloneTable
Dim pExportsTable_S As IStandaloneTable
Dim pLoadPrecursorTable_S As IStandaloneTable
Dim pPerCapitaTPLoads_S As IStandaloneTable
Dim pScenario As IStandaloneTable
Dim pTPBreakpoints_S As IStandaloneTable
Dim blnUseModifiedLandCoverSlope As Boolean
Dim pTableUserDefinedLCoverSlope_Summary As ITable
Dim varModifiedLCoverCoeff() As Variant 'used to store user modified land cover coefficients (also used during program execution)

'note that the PLUS+ User Guide gives instructions for the migration to 10.2 - the line below gives an error that a user-defined type is not defined
'Ticking the References box "Microsoft Windows Common Controls 6.0 (SP6)" will fix the user-defined type error
Dim CatchmentInfoColumnHeaders1 As columnHeader
Dim CatchmentRelateColumnHeaders1 As columnHeader
Dim CatchmentRelateColumnHeaders2 As columnHeader
Dim SupplementColumnHeaders As columnHeader
Dim List_Item1 As ListItem
Dim List_Item2 As ListItem
Dim List_Item3 As ListItem
Dim List_ItemSupplement As ListItem
Dim strTPBreakPointsRefType As String
Dim intTopListviewBoxesTop As Integer
Dim intTopListviewBoxesHeight As Integer
Dim dblUser_Entered_P_for_Selected_Site As Double
Dim dblUser_Entered_Area_for_Selected_Site As Double
Dim strSelectedLandCoverType As String
Dim lontblSelectedCatchCatchPRecords As Long
Dim dblSumLocalInputs As Double
Dim dblSumUpstream As Double
Dim dblUserModifiedLCoverArea_difference As Double
Dim blnResolveDifferences As Boolean
Dim strNewLCoverToUse As String
Dim blnUseLCoverComboSelection As Boolean
Dim dblSumLandCover As Double
'variables for user created file intersection
Dim pFLayerUserLandCoverInput As IFeatureLayer2
Dim pFClassUserFile As IFeatureClass
Dim pFLayerSlopeLandClass As IFeatureLayer2 'added this option in case the user wishes to use a
Dim pFClassSlopeLandClass As IFeatureClass  ' different file
Dim strSelectedPointSourceForRemoval As String

'variables for point source scenario use
Dim strScenarioTblPointSource As String
Dim pPointSourceTable_S As IStandaloneTable
Dim intSPS_GBLAKES_ID_field As Integer
Dim intSPS_Type_field As Integer
Dim intSPS_Amount_field As Integer
Dim intSPS_ScenarioID_field As Integer

Dim pGraphicsContainer As IGraphicsContainer
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub RunBatchProcesses()
'this is a procedure to run the whole database, exporting to text
'enable this in 'Activate'

'##################
'Critical: compact the source data mdb before running this. After a run compact it again. If it is running slowly interrupt the process
'and compact the mdb again - this is key to fast processing.
'A file geodatabase does not gain the bloat that the mdb but it always runs more slowly. The mdb is approx. 5 times quicker. I've no idea why.
'##################

'##################
'for speed also disable the line 'MsgBox "Site " & CatchNetRship(i, 0) & " contains NULL values for local runoff or local area and so cannot be processed.", vbCritical
'and the line MsgBox "A text file, " & strFileName & ", has been saved to your output folder which describes discrepancies between the SEPA and PLUS+ modelled status"
'#################

Dim runStandardLoads As Boolean
'##################  U S E R   S E T T I N G  ##################
'option to overwrite the per capita P loads from the PerCapitaTPLoads table
'further settings are in the code below the array definitions - search for "runStandardLoads"
runStandardLoads = True
'runStandardLoads = False

'##################  U S E R   S E T T I N G  ##################
'option to load from a predefined table of point sources, e.g. birds and WWTW
'further settings are in the code below to change per capita P settings in conjunction with the runStandardLoads option above
'chkLoadPointSources = True
chkLoadPointSources = False

Dim blnFound As Boolean
Dim fileCheckExists As String
'##################  U S E R   S E T T I N G  ##################
'this section allows you to choose whether to overwrite reports or only process missing one, comment as appropriate.
'fileCheckExists = "always process"
fileCheckExists = "process if missing"

Dim strOutputFolder As String
'##################  U S E R   S E T T I N G  ##################
'specifiy where the text reports will go - any in here will be overwritten (depending on the above setting)
strOutputFolder = "D:\SmallProjects\PLUS2\BatchOut\Threshold_230\"

'##################  U S E R   S E T T I N G  ##################
'choose which year of classification is to be run
'cboSEPAmonitoring.Text = "SEPA_detailed_Loch_WB_classification"
'cboClassConcStat.Text = "SEPA_2009_loch_class"
cboClassConcStat.Text = "SEPA_2014_loch_class"
cboSEPAmonitoring.Text = "SEPA_2014_detailed_Loch_WB_classification"

Dim intArrayToUse As Integer
'##################  U S E R   S E T T I N G  ##################
'there are four arrays preconfigured, set the variable intArrayToUse to 1, 2, 3 or 4 as required:
'1. To match only the SEPA measured lochs (2009 set - the 2014 is slightly different)
'2. The top half of the minimum group to run the whole country
'3. The bottom half of the minimum group to run the whole country
'4. To match only the SEPA measured lochs (2014 set - the 2009 is slightly different)
intArrayToUse = 4
'##################  E N D   O F   U S E R   S E T T I N G S ##################

Dim arrayCatchToRun() As Integer
'Dim arrayCatchToRun() As Long   'use this one for option 4

If intArrayToUse = 1 Then
    ReDim arrayCatchToRun(319) As Integer
    arrayCatchToRun(0) = 7
    arrayCatchToRun(1) = 704
    arrayCatchToRun(2) = 1135
    arrayCatchToRun(3) = 1271
    arrayCatchToRun(4) = 1570
    arrayCatchToRun(5) = 1678
    arrayCatchToRun(6) = 1692
    arrayCatchToRun(7) = 1694
    arrayCatchToRun(8) = 1753
    arrayCatchToRun(9) = 1764
    arrayCatchToRun(10) = 1847
    arrayCatchToRun(11) = 1853
    arrayCatchToRun(12) = 2035
    arrayCatchToRun(13) = 2098
    arrayCatchToRun(14) = 2139
    arrayCatchToRun(15) = 2237
    arrayCatchToRun(16) = 2358
    arrayCatchToRun(17) = 2402
    arrayCatchToRun(18) = 2490
    arrayCatchToRun(19) = 2499
    arrayCatchToRun(20) = 2580
    arrayCatchToRun(21) = 2712
    arrayCatchToRun(22) = 2831
    arrayCatchToRun(23) = 3105
    arrayCatchToRun(24) = 3458
    arrayCatchToRun(25) = 3527
    arrayCatchToRun(26) = 3532
    arrayCatchToRun(27) = 3636
    arrayCatchToRun(28) = 3904
    arrayCatchToRun(29) = 3930
    arrayCatchToRun(30) = 4164
    arrayCatchToRun(31) = 4229
    arrayCatchToRun(32) = 4284
    arrayCatchToRun(33) = 4444
    arrayCatchToRun(34) = 4672
    arrayCatchToRun(35) = 4796
    arrayCatchToRun(36) = 5222
    arrayCatchToRun(37) = 5284
    arrayCatchToRun(38) = 5350
    arrayCatchToRun(39) = 5605
    arrayCatchToRun(40) = 6140
    arrayCatchToRun(41) = 6297
    arrayCatchToRun(42) = 6405
    arrayCatchToRun(43) = 6455
    arrayCatchToRun(44) = 6687
    arrayCatchToRun(45) = 7092
    arrayCatchToRun(46) = 7183
    arrayCatchToRun(47) = 7222
    arrayCatchToRun(48) = 7469
    arrayCatchToRun(49) = 7588
    arrayCatchToRun(50) = 7691
    arrayCatchToRun(51) = 7730
    arrayCatchToRun(52) = 8090
    arrayCatchToRun(53) = 8097
    arrayCatchToRun(54) = 8168
    arrayCatchToRun(55) = 8307
    arrayCatchToRun(56) = 8361
    arrayCatchToRun(57) = 8455
    arrayCatchToRun(58) = 8564
    arrayCatchToRun(59) = 8738
    arrayCatchToRun(60) = 8751
    arrayCatchToRun(61) = 8937
    arrayCatchToRun(62) = 8959
    arrayCatchToRun(63) = 8971
    arrayCatchToRun(64) = 9048
    arrayCatchToRun(65) = 9098
    arrayCatchToRun(66) = 9176
    arrayCatchToRun(67) = 9281
    arrayCatchToRun(68) = 9351
    arrayCatchToRun(69) = 9401
    arrayCatchToRun(70) = 9617
    arrayCatchToRun(71) = 9658
    arrayCatchToRun(72) = 9852
    arrayCatchToRun(73) = 9922
    arrayCatchToRun(74) = 10060
    arrayCatchToRun(75) = 10132
    arrayCatchToRun(76) = 10301
    arrayCatchToRun(77) = 10389
    arrayCatchToRun(78) = 10719
    arrayCatchToRun(79) = 10786
    arrayCatchToRun(80) = 10934
    arrayCatchToRun(81) = 11101
    arrayCatchToRun(82) = 11187
    arrayCatchToRun(83) = 11189
    arrayCatchToRun(84) = 11337
    arrayCatchToRun(85) = 11338
    arrayCatchToRun(86) = 11339
    arrayCatchToRun(87) = 11384
    arrayCatchToRun(88) = 11385
    arrayCatchToRun(89) = 11447
    arrayCatchToRun(90) = 11504
    arrayCatchToRun(91) = 11611
    arrayCatchToRun(92) = 11642
    arrayCatchToRun(93) = 12055
    arrayCatchToRun(94) = 12313
    arrayCatchToRun(95) = 12382
    arrayCatchToRun(96) = 12606
    arrayCatchToRun(97) = 12659
    arrayCatchToRun(98) = 12978
    arrayCatchToRun(99) = 12995
    arrayCatchToRun(100) = 13445
    arrayCatchToRun(101) = 13463
    arrayCatchToRun(102) = 13544
    arrayCatchToRun(103) = 13696
    arrayCatchToRun(104) = 13780
    arrayCatchToRun(105) = 13972
    arrayCatchToRun(106) = 14019
    arrayCatchToRun(107) = 14032
    arrayCatchToRun(108) = 14057
    arrayCatchToRun(109) = 14098
    arrayCatchToRun(110) = 14202
    arrayCatchToRun(111) = 14293
    arrayCatchToRun(112) = 14315
    arrayCatchToRun(113) = 14362
    arrayCatchToRun(114) = 14384
    arrayCatchToRun(115) = 14443
    arrayCatchToRun(116) = 14585
    arrayCatchToRun(117) = 14627
    arrayCatchToRun(118) = 14740
    arrayCatchToRun(119) = 14749
    arrayCatchToRun(120) = 15026
    arrayCatchToRun(121) = 15139
    arrayCatchToRun(122) = 15222
    arrayCatchToRun(123) = 15265
    arrayCatchToRun(124) = 15390
    arrayCatchToRun(125) = 15477
    arrayCatchToRun(126) = 15651
    arrayCatchToRun(127) = 15831
    arrayCatchToRun(128) = 15935
    arrayCatchToRun(129) = 16206
    arrayCatchToRun(130) = 16220
    arrayCatchToRun(131) = 16275
    arrayCatchToRun(132) = 16369
    arrayCatchToRun(133) = 16443
    arrayCatchToRun(134) = 16456
    arrayCatchToRun(135) = 16505
    arrayCatchToRun(136) = 16624
    arrayCatchToRun(137) = 16661
    arrayCatchToRun(138) = 16902
    arrayCatchToRun(139) = 16906
    arrayCatchToRun(140) = 16986
    arrayCatchToRun(141) = 17086
    arrayCatchToRun(142) = 17239
    arrayCatchToRun(143) = 17241
    arrayCatchToRun(144) = 17619
    arrayCatchToRun(145) = 18216
    arrayCatchToRun(146) = 18607
    arrayCatchToRun(147) = 18644
    arrayCatchToRun(148) = 18645
    arrayCatchToRun(149) = 18682
    arrayCatchToRun(150) = 18767
    arrayCatchToRun(151) = 18825
    arrayCatchToRun(152) = 18876
    arrayCatchToRun(153) = 18982
    arrayCatchToRun(154) = 19079
    arrayCatchToRun(155) = 19214
    arrayCatchToRun(156) = 19261
    arrayCatchToRun(157) = 19283
    arrayCatchToRun(158) = 19381
    arrayCatchToRun(159) = 19445
    arrayCatchToRun(160) = 19516
    arrayCatchToRun(161) = 19540
    arrayCatchToRun(162) = 19572
    arrayCatchToRun(163) = 19864
    arrayCatchToRun(164) = 19896
    arrayCatchToRun(165) = 19935
    arrayCatchToRun(166) = 19952
    arrayCatchToRun(167) = 20002
    arrayCatchToRun(168) = 20043
    arrayCatchToRun(169) = 20185
    arrayCatchToRun(170) = 20196
    arrayCatchToRun(171) = 20465
    arrayCatchToRun(172) = 20573
    arrayCatchToRun(173) = 20601
    arrayCatchToRun(174) = 20633
    arrayCatchToRun(175) = 20647
    arrayCatchToRun(176) = 20657
    arrayCatchToRun(177) = 20742
    arrayCatchToRun(178) = 20754
    arrayCatchToRun(179) = 20757
    arrayCatchToRun(180) = 20828
    arrayCatchToRun(181) = 20860
    arrayCatchToRun(182) = 20965
    arrayCatchToRun(183) = 20986
    arrayCatchToRun(184) = 21023
    arrayCatchToRun(185) = 21152
    arrayCatchToRun(186) = 21189
    arrayCatchToRun(187) = 21191
    arrayCatchToRun(188) = 21328
    arrayCatchToRun(189) = 21437
    arrayCatchToRun(190) = 21490
    arrayCatchToRun(191) = 21576
    arrayCatchToRun(192) = 21649
    arrayCatchToRun(193) = 21751
    arrayCatchToRun(194) = 21754
    arrayCatchToRun(195) = 21790
    arrayCatchToRun(196) = 21795
    arrayCatchToRun(197) = 21823
    arrayCatchToRun(198) = 21847
    arrayCatchToRun(199) = 21848
    arrayCatchToRun(200) = 21877
    arrayCatchToRun(201) = 21925
    arrayCatchToRun(202) = 21945
    arrayCatchToRun(203) = 21956
    arrayCatchToRun(204) = 22010
    arrayCatchToRun(205) = 22191
    arrayCatchToRun(206) = 22308
    arrayCatchToRun(207) = 22419
    arrayCatchToRun(208) = 22496
    arrayCatchToRun(209) = 22610
    arrayCatchToRun(210) = 22666
    arrayCatchToRun(211) = 22725
    arrayCatchToRun(212) = 22782
    arrayCatchToRun(213) = 22787
    arrayCatchToRun(214) = 22839
    arrayCatchToRun(215) = 22840
    arrayCatchToRun(216) = 22942
    arrayCatchToRun(217) = 23206
    arrayCatchToRun(218) = 23216
    arrayCatchToRun(219) = 23245
    arrayCatchToRun(220) = 23361
    arrayCatchToRun(221) = 23465
    arrayCatchToRun(222) = 23553
    arrayCatchToRun(223) = 23559
    arrayCatchToRun(224) = 23561
    arrayCatchToRun(225) = 23578
    arrayCatchToRun(226) = 23624
    arrayCatchToRun(227) = 23654
    arrayCatchToRun(228) = 23684
    arrayCatchToRun(229) = 23711
    arrayCatchToRun(230) = 23887
    arrayCatchToRun(231) = 23938
    arrayCatchToRun(232) = 23973
    arrayCatchToRun(233) = 24016
    arrayCatchToRun(234) = 24103
    arrayCatchToRun(235) = 24124
    arrayCatchToRun(236) = 24132
    arrayCatchToRun(237) = 24276
    arrayCatchToRun(238) = 24280
    arrayCatchToRun(239) = 24295
    arrayCatchToRun(240) = 24297
    arrayCatchToRun(241) = 24417
    arrayCatchToRun(242) = 24464
    arrayCatchToRun(243) = 24522
    arrayCatchToRun(244) = 24668
    arrayCatchToRun(245) = 24744
    arrayCatchToRun(246) = 24754
    arrayCatchToRun(247) = 24758
    arrayCatchToRun(248) = 24785
    arrayCatchToRun(249) = 24798
    arrayCatchToRun(250) = 24843
    arrayCatchToRun(251) = 24892
    arrayCatchToRun(252) = 24919
    arrayCatchToRun(253) = 24996
    arrayCatchToRun(254) = 25000
    arrayCatchToRun(255) = 25006
    arrayCatchToRun(256) = 25035
    arrayCatchToRun(257) = 25038
    arrayCatchToRun(258) = 25077
    arrayCatchToRun(259) = 25128
    arrayCatchToRun(260) = 25366
    arrayCatchToRun(261) = 25378
    arrayCatchToRun(262) = 25391
    arrayCatchToRun(263) = 25400
    arrayCatchToRun(264) = 25889
    arrayCatchToRun(265) = 25925
    arrayCatchToRun(266) = 26101
    arrayCatchToRun(267) = 26114
    arrayCatchToRun(268) = 26168
    arrayCatchToRun(269) = 26237
    arrayCatchToRun(270) = 26240
    arrayCatchToRun(271) = 26243
    arrayCatchToRun(272) = 26275
    arrayCatchToRun(273) = 26356
    arrayCatchToRun(274) = 26392
    arrayCatchToRun(275) = 26416
    arrayCatchToRun(276) = 26447
    arrayCatchToRun(277) = 26450
    arrayCatchToRun(278) = 26467
    arrayCatchToRun(279) = 26472
    arrayCatchToRun(280) = 26566
    arrayCatchToRun(281) = 26578
    arrayCatchToRun(282) = 26581
    arrayCatchToRun(283) = 26692
    arrayCatchToRun(284) = 26752
    arrayCatchToRun(285) = 26804
    arrayCatchToRun(286) = 27186
    arrayCatchToRun(287) = 27309
    arrayCatchToRun(288) = 27310
    arrayCatchToRun(289) = 27315
    arrayCatchToRun(290) = 27322
    arrayCatchToRun(291) = 27361
    arrayCatchToRun(292) = 27435
    arrayCatchToRun(293) = 27523
    arrayCatchToRun(294) = 27604
    arrayCatchToRun(295) = 27627
    arrayCatchToRun(296) = 27638
    arrayCatchToRun(297) = 27675
    arrayCatchToRun(298) = 27699
    arrayCatchToRun(299) = 27785
    arrayCatchToRun(300) = 27808
    arrayCatchToRun(301) = 27848
    arrayCatchToRun(302) = 27899
    arrayCatchToRun(303) = 27936
    arrayCatchToRun(304) = 27948
    arrayCatchToRun(305) = 27967
    arrayCatchToRun(306) = 28003
    arrayCatchToRun(307) = 28014
    arrayCatchToRun(308) = 28043
    arrayCatchToRun(309) = 28071
    arrayCatchToRun(310) = 28111
    arrayCatchToRun(311) = 28130
    arrayCatchToRun(312) = 28158
    arrayCatchToRun(313) = 28200
    arrayCatchToRun(314) = 28288
    arrayCatchToRun(315) = 28330
    arrayCatchToRun(316) = 28344
    arrayCatchToRun(317) = 28493
    arrayCatchToRun(318) = 28506
End If

If intArrayToUse = 2 Then
    ReDim arrayCatchToRun(1952) As Integer
    arrayCatchToRun(0) = 30
    arrayCatchToRun(1) = 37
    arrayCatchToRun(2) = 45
    arrayCatchToRun(3) = 60
    arrayCatchToRun(4) = 64
    arrayCatchToRun(5) = 69
    arrayCatchToRun(6) = 83
    arrayCatchToRun(7) = 90
    arrayCatchToRun(8) = 93
    arrayCatchToRun(9) = 102
    arrayCatchToRun(10) = 109
    arrayCatchToRun(11) = 123
    arrayCatchToRun(12) = 125
    arrayCatchToRun(13) = 140
    arrayCatchToRun(14) = 150
    arrayCatchToRun(15) = 159
    arrayCatchToRun(16) = 160
    arrayCatchToRun(17) = 163
    arrayCatchToRun(18) = 172
    arrayCatchToRun(19) = 175
    arrayCatchToRun(20) = 179
    arrayCatchToRun(21) = 207
    arrayCatchToRun(22) = 210
    arrayCatchToRun(23) = 212
    arrayCatchToRun(24) = 215
    arrayCatchToRun(25) = 216
    arrayCatchToRun(26) = 221
    arrayCatchToRun(27) = 223
    arrayCatchToRun(28) = 230
    arrayCatchToRun(29) = 237
    arrayCatchToRun(30) = 239
    arrayCatchToRun(31) = 263
    arrayCatchToRun(32) = 265
    arrayCatchToRun(33) = 272
    arrayCatchToRun(34) = 275
    arrayCatchToRun(35) = 276
    arrayCatchToRun(36) = 281
    arrayCatchToRun(37) = 285
    arrayCatchToRun(38) = 287
    arrayCatchToRun(39) = 294
    arrayCatchToRun(40) = 296
    arrayCatchToRun(41) = 308
    arrayCatchToRun(42) = 313
    arrayCatchToRun(43) = 323
    arrayCatchToRun(44) = 330
    arrayCatchToRun(45) = 355
    arrayCatchToRun(46) = 449
    arrayCatchToRun(47) = 483
    arrayCatchToRun(48) = 486
    arrayCatchToRun(49) = 565
    arrayCatchToRun(50) = 567
    arrayCatchToRun(51) = 579
    arrayCatchToRun(52) = 587
    arrayCatchToRun(53) = 593
    arrayCatchToRun(54) = 594
    arrayCatchToRun(55) = 600
    arrayCatchToRun(56) = 611
    arrayCatchToRun(57) = 615
    arrayCatchToRun(58) = 617
    arrayCatchToRun(59) = 629
    arrayCatchToRun(60) = 643
    arrayCatchToRun(61) = 650
    arrayCatchToRun(62) = 653
    arrayCatchToRun(63) = 654
    arrayCatchToRun(64) = 657
    arrayCatchToRun(65) = 665
    arrayCatchToRun(66) = 666
    arrayCatchToRun(67) = 668
    arrayCatchToRun(68) = 683
    arrayCatchToRun(69) = 688
    arrayCatchToRun(70) = 701
    arrayCatchToRun(71) = 702
    arrayCatchToRun(72) = 707
    arrayCatchToRun(73) = 709
    arrayCatchToRun(74) = 716
    arrayCatchToRun(75) = 717
    arrayCatchToRun(76) = 723
    arrayCatchToRun(77) = 724
    arrayCatchToRun(78) = 730
    arrayCatchToRun(79) = 731
    arrayCatchToRun(80) = 733
    arrayCatchToRun(81) = 740
    arrayCatchToRun(82) = 744
    arrayCatchToRun(83) = 746
    arrayCatchToRun(84) = 757
    arrayCatchToRun(85) = 759
    arrayCatchToRun(86) = 762
    arrayCatchToRun(87) = 767
    arrayCatchToRun(88) = 770
    arrayCatchToRun(89) = 773
    arrayCatchToRun(90) = 785
    arrayCatchToRun(91) = 804
    arrayCatchToRun(92) = 809
    arrayCatchToRun(93) = 811
    arrayCatchToRun(94) = 812
    arrayCatchToRun(95) = 830
    arrayCatchToRun(96) = 838
    arrayCatchToRun(97) = 840
    arrayCatchToRun(98) = 842
    arrayCatchToRun(99) = 850
    arrayCatchToRun(100) = 856
    arrayCatchToRun(101) = 858
    arrayCatchToRun(102) = 871
    arrayCatchToRun(103) = 880
    arrayCatchToRun(104) = 886
    arrayCatchToRun(105) = 890
    arrayCatchToRun(106) = 892
    arrayCatchToRun(107) = 911
    arrayCatchToRun(108) = 914
    arrayCatchToRun(109) = 915
    arrayCatchToRun(110) = 920
    arrayCatchToRun(111) = 924
    arrayCatchToRun(112) = 931
    arrayCatchToRun(113) = 934
    arrayCatchToRun(114) = 935
    arrayCatchToRun(115) = 936
    arrayCatchToRun(116) = 937
    arrayCatchToRun(117) = 940
    arrayCatchToRun(118) = 941
    arrayCatchToRun(119) = 943
    arrayCatchToRun(120) = 945
    arrayCatchToRun(121) = 947
    arrayCatchToRun(122) = 952
    arrayCatchToRun(123) = 954
    arrayCatchToRun(124) = 959
    arrayCatchToRun(125) = 962
    arrayCatchToRun(126) = 966
    arrayCatchToRun(127) = 971
    arrayCatchToRun(128) = 985
    arrayCatchToRun(129) = 986
    arrayCatchToRun(130) = 996
    arrayCatchToRun(131) = 997
    arrayCatchToRun(132) = 999
    arrayCatchToRun(133) = 1007
    arrayCatchToRun(134) = 1008
    arrayCatchToRun(135) = 1015
    arrayCatchToRun(136) = 1018
    arrayCatchToRun(137) = 1022
    arrayCatchToRun(138) = 1026
    arrayCatchToRun(139) = 1031
    arrayCatchToRun(140) = 1037
    arrayCatchToRun(141) = 1046
    arrayCatchToRun(142) = 1048
    arrayCatchToRun(143) = 1053
    arrayCatchToRun(144) = 1060
    arrayCatchToRun(145) = 1085
    arrayCatchToRun(146) = 1086
    arrayCatchToRun(147) = 1093
    arrayCatchToRun(148) = 1104
    arrayCatchToRun(149) = 1107
    arrayCatchToRun(150) = 1114
    arrayCatchToRun(151) = 1119
    arrayCatchToRun(152) = 1122
    arrayCatchToRun(153) = 1130
    arrayCatchToRun(154) = 1140
    arrayCatchToRun(155) = 1147
    arrayCatchToRun(156) = 1151
    arrayCatchToRun(157) = 1152
    arrayCatchToRun(158) = 1159
    arrayCatchToRun(159) = 1161
    arrayCatchToRun(160) = 1170
    arrayCatchToRun(161) = 1192
    arrayCatchToRun(162) = 1195
    arrayCatchToRun(163) = 1199
    arrayCatchToRun(164) = 1200
    arrayCatchToRun(165) = 1212
    arrayCatchToRun(166) = 1220
    arrayCatchToRun(167) = 1221
    arrayCatchToRun(168) = 1244
    arrayCatchToRun(169) = 1288
    arrayCatchToRun(170) = 1290
    arrayCatchToRun(171) = 1310
    arrayCatchToRun(172) = 1313
    arrayCatchToRun(173) = 1321
    arrayCatchToRun(174) = 1322
    arrayCatchToRun(175) = 1327
    arrayCatchToRun(176) = 1329
    arrayCatchToRun(177) = 1331
    arrayCatchToRun(178) = 1335
    arrayCatchToRun(179) = 1340
    arrayCatchToRun(180) = 1343
    arrayCatchToRun(181) = 1347
    arrayCatchToRun(182) = 1354
    arrayCatchToRun(183) = 1366
    arrayCatchToRun(184) = 1379
    arrayCatchToRun(185) = 1384
    arrayCatchToRun(186) = 1394
    arrayCatchToRun(187) = 1395
    arrayCatchToRun(188) = 1396
    arrayCatchToRun(189) = 1398
    arrayCatchToRun(190) = 1404
    arrayCatchToRun(191) = 1410
    arrayCatchToRun(192) = 1428
    arrayCatchToRun(193) = 1442
    arrayCatchToRun(194) = 1447
    arrayCatchToRun(195) = 1448
    arrayCatchToRun(196) = 1452
    arrayCatchToRun(197) = 1459
    arrayCatchToRun(198) = 1462
    arrayCatchToRun(199) = 1465
    arrayCatchToRun(200) = 1472
    arrayCatchToRun(201) = 1488
    arrayCatchToRun(202) = 1489
    arrayCatchToRun(203) = 1493
    arrayCatchToRun(204) = 1494
    arrayCatchToRun(205) = 1507
    arrayCatchToRun(206) = 1508
    arrayCatchToRun(207) = 1509
    arrayCatchToRun(208) = 1517
    arrayCatchToRun(209) = 1535
    arrayCatchToRun(210) = 1536
    arrayCatchToRun(211) = 1538
    arrayCatchToRun(212) = 1539
    arrayCatchToRun(213) = 1541
    arrayCatchToRun(214) = 1544
    arrayCatchToRun(215) = 1545
    arrayCatchToRun(216) = 1563
    arrayCatchToRun(217) = 1567
    arrayCatchToRun(218) = 1571
    arrayCatchToRun(219) = 1572
    arrayCatchToRun(220) = 1582
    arrayCatchToRun(221) = 1588
    arrayCatchToRun(222) = 1597
    arrayCatchToRun(223) = 1602
    arrayCatchToRun(224) = 1606
    arrayCatchToRun(225) = 1608
    arrayCatchToRun(226) = 1617
    arrayCatchToRun(227) = 1620
    arrayCatchToRun(228) = 1624
    arrayCatchToRun(229) = 1628
    arrayCatchToRun(230) = 1633
    arrayCatchToRun(231) = 1646
    arrayCatchToRun(232) = 1656
    arrayCatchToRun(233) = 1660
    arrayCatchToRun(234) = 1666
    arrayCatchToRun(235) = 1669
    arrayCatchToRun(236) = 1674
    arrayCatchToRun(237) = 1675
    arrayCatchToRun(238) = 1676
    arrayCatchToRun(239) = 1677
    arrayCatchToRun(240) = 1678
    arrayCatchToRun(241) = 1682
    arrayCatchToRun(242) = 1685
    arrayCatchToRun(243) = 1688
    arrayCatchToRun(244) = 1699
    arrayCatchToRun(245) = 1709
    arrayCatchToRun(246) = 1717
    arrayCatchToRun(247) = 1729
    arrayCatchToRun(248) = 1741
    arrayCatchToRun(249) = 1743
    arrayCatchToRun(250) = 1758
    arrayCatchToRun(251) = 1759
    arrayCatchToRun(252) = 1764
    arrayCatchToRun(253) = 1765
    arrayCatchToRun(254) = 1807
    arrayCatchToRun(255) = 1832
    arrayCatchToRun(256) = 1836
    arrayCatchToRun(257) = 1840
    arrayCatchToRun(258) = 1841
    arrayCatchToRun(259) = 1845
    arrayCatchToRun(260) = 1847
    arrayCatchToRun(261) = 1853
    arrayCatchToRun(262) = 1857
    arrayCatchToRun(263) = 1860
    arrayCatchToRun(264) = 1869
    arrayCatchToRun(265) = 1887
    arrayCatchToRun(266) = 1901
    arrayCatchToRun(267) = 1905
    arrayCatchToRun(268) = 1910
    arrayCatchToRun(269) = 1913
    arrayCatchToRun(270) = 1936
    arrayCatchToRun(271) = 1987
    arrayCatchToRun(272) = 2016
    arrayCatchToRun(273) = 2029
    arrayCatchToRun(274) = 2031
    arrayCatchToRun(275) = 2035
    arrayCatchToRun(276) = 2038
    arrayCatchToRun(277) = 2039
    arrayCatchToRun(278) = 2058
    arrayCatchToRun(279) = 2062
    arrayCatchToRun(280) = 2088
    arrayCatchToRun(281) = 2094
    arrayCatchToRun(282) = 2097
    arrayCatchToRun(283) = 2098
    arrayCatchToRun(284) = 2109
    arrayCatchToRun(285) = 2118
    arrayCatchToRun(286) = 2124
    arrayCatchToRun(287) = 2125
    arrayCatchToRun(288) = 2126
    arrayCatchToRun(289) = 2139
    arrayCatchToRun(290) = 2143
    arrayCatchToRun(291) = 2149
    arrayCatchToRun(292) = 2154
    arrayCatchToRun(293) = 2161
    arrayCatchToRun(294) = 2165
    arrayCatchToRun(295) = 2176
    arrayCatchToRun(296) = 2178
    arrayCatchToRun(297) = 2179
    arrayCatchToRun(298) = 2185
    arrayCatchToRun(299) = 2191
    arrayCatchToRun(300) = 2194
    arrayCatchToRun(301) = 2200
    arrayCatchToRun(302) = 2206
    arrayCatchToRun(303) = 2208
    arrayCatchToRun(304) = 2216
    arrayCatchToRun(305) = 2225
    arrayCatchToRun(306) = 2229
    arrayCatchToRun(307) = 2231
    arrayCatchToRun(308) = 2234
    arrayCatchToRun(309) = 2242
    arrayCatchToRun(310) = 2250
    arrayCatchToRun(311) = 2253
    arrayCatchToRun(312) = 2264
    arrayCatchToRun(313) = 2270
    arrayCatchToRun(314) = 2272
    arrayCatchToRun(315) = 2275
    arrayCatchToRun(316) = 2277
    arrayCatchToRun(317) = 2278
    arrayCatchToRun(318) = 2280
    arrayCatchToRun(319) = 2285
    arrayCatchToRun(320) = 2288
    arrayCatchToRun(321) = 2290
    arrayCatchToRun(322) = 2291
    arrayCatchToRun(323) = 2296
    arrayCatchToRun(324) = 2298
    arrayCatchToRun(325) = 2299
    arrayCatchToRun(326) = 2309
    arrayCatchToRun(327) = 2315
    arrayCatchToRun(328) = 2326
    arrayCatchToRun(329) = 2333
    arrayCatchToRun(330) = 2347
    arrayCatchToRun(331) = 2348
    arrayCatchToRun(332) = 2352
    arrayCatchToRun(333) = 2353
    arrayCatchToRun(334) = 2357
    arrayCatchToRun(335) = 2359
    arrayCatchToRun(336) = 2360
    arrayCatchToRun(337) = 2362
    arrayCatchToRun(338) = 2364
    arrayCatchToRun(339) = 2365
    arrayCatchToRun(340) = 2392
    arrayCatchToRun(341) = 2393
    arrayCatchToRun(342) = 2395
    arrayCatchToRun(343) = 2398
    arrayCatchToRun(344) = 2401
    arrayCatchToRun(345) = 2417
    arrayCatchToRun(346) = 2418
    arrayCatchToRun(347) = 2421
    arrayCatchToRun(348) = 2422
    arrayCatchToRun(349) = 2428
    arrayCatchToRun(350) = 2430
    arrayCatchToRun(351) = 2431
    arrayCatchToRun(352) = 2442
    arrayCatchToRun(353) = 2443
    arrayCatchToRun(354) = 2446
    arrayCatchToRun(355) = 2454
    arrayCatchToRun(356) = 2456
    arrayCatchToRun(357) = 2474
    arrayCatchToRun(358) = 2476
    arrayCatchToRun(359) = 2481
    arrayCatchToRun(360) = 2484
    arrayCatchToRun(361) = 2489
    arrayCatchToRun(362) = 2495
    arrayCatchToRun(363) = 2498
    arrayCatchToRun(364) = 2504
    arrayCatchToRun(365) = 2508
    arrayCatchToRun(366) = 2512
    arrayCatchToRun(367) = 2518
    arrayCatchToRun(368) = 2524
    arrayCatchToRun(369) = 2527
    arrayCatchToRun(370) = 2554
    arrayCatchToRun(371) = 2556
    arrayCatchToRun(372) = 2559
    arrayCatchToRun(373) = 2560
    arrayCatchToRun(374) = 2562
    arrayCatchToRun(375) = 2566
    arrayCatchToRun(376) = 2588
    arrayCatchToRun(377) = 2596
    arrayCatchToRun(378) = 2602
    arrayCatchToRun(379) = 2604
    arrayCatchToRun(380) = 2607
    arrayCatchToRun(381) = 2614
    arrayCatchToRun(382) = 2630
    arrayCatchToRun(383) = 2634
    arrayCatchToRun(384) = 2651
    arrayCatchToRun(385) = 2653
    arrayCatchToRun(386) = 2667
    arrayCatchToRun(387) = 2675
    arrayCatchToRun(388) = 2677
    arrayCatchToRun(389) = 2683
    arrayCatchToRun(390) = 2686
    arrayCatchToRun(391) = 2688
    arrayCatchToRun(392) = 2696
    arrayCatchToRun(393) = 2732
    arrayCatchToRun(394) = 2742
    arrayCatchToRun(395) = 2743
    arrayCatchToRun(396) = 2750
    arrayCatchToRun(397) = 2751
    arrayCatchToRun(398) = 2769
    arrayCatchToRun(399) = 2779
    arrayCatchToRun(400) = 2791
    arrayCatchToRun(401) = 2796
    arrayCatchToRun(402) = 2836
    arrayCatchToRun(403) = 2870
    arrayCatchToRun(404) = 2871
    arrayCatchToRun(405) = 2874
    arrayCatchToRun(406) = 2906
    arrayCatchToRun(407) = 2939
    arrayCatchToRun(408) = 2978
    arrayCatchToRun(409) = 3011
    arrayCatchToRun(410) = 3014
    arrayCatchToRun(411) = 3020
    arrayCatchToRun(412) = 3034
    arrayCatchToRun(413) = 3063
    arrayCatchToRun(414) = 3068
    arrayCatchToRun(415) = 3073
    arrayCatchToRun(416) = 3087
    arrayCatchToRun(417) = 3094
    arrayCatchToRun(418) = 3129
    arrayCatchToRun(419) = 3139
    arrayCatchToRun(420) = 3150
    arrayCatchToRun(421) = 3164
    arrayCatchToRun(422) = 3166
    arrayCatchToRun(423) = 3178
    arrayCatchToRun(424) = 3185
    arrayCatchToRun(425) = 3194
    arrayCatchToRun(426) = 3198
    arrayCatchToRun(427) = 3201
    arrayCatchToRun(428) = 3205
    arrayCatchToRun(429) = 3206
    arrayCatchToRun(430) = 3216
    arrayCatchToRun(431) = 3249
    arrayCatchToRun(432) = 3270
    arrayCatchToRun(433) = 3283
    arrayCatchToRun(434) = 3291
    arrayCatchToRun(435) = 3296
    arrayCatchToRun(436) = 3299
    arrayCatchToRun(437) = 3312
    arrayCatchToRun(438) = 3313
    arrayCatchToRun(439) = 3318
    arrayCatchToRun(440) = 3351
    arrayCatchToRun(441) = 3394
    arrayCatchToRun(442) = 3429
    arrayCatchToRun(443) = 3431
    arrayCatchToRun(444) = 3436
    arrayCatchToRun(445) = 3445
    arrayCatchToRun(446) = 3464
    arrayCatchToRun(447) = 3470
    arrayCatchToRun(448) = 3483
    arrayCatchToRun(449) = 3492
    arrayCatchToRun(450) = 3532
    arrayCatchToRun(451) = 3548
    arrayCatchToRun(452) = 3585
    arrayCatchToRun(453) = 3601
    arrayCatchToRun(454) = 3616
    arrayCatchToRun(455) = 3641
    arrayCatchToRun(456) = 3644
    arrayCatchToRun(457) = 3668
    arrayCatchToRun(458) = 3743
    arrayCatchToRun(459) = 3747
    arrayCatchToRun(460) = 3756
    arrayCatchToRun(461) = 3761
    arrayCatchToRun(462) = 3765
    arrayCatchToRun(463) = 3805
    arrayCatchToRun(464) = 3822
    arrayCatchToRun(465) = 3852
    arrayCatchToRun(466) = 3856
    arrayCatchToRun(467) = 3882
    arrayCatchToRun(468) = 3915
    arrayCatchToRun(469) = 3951
    arrayCatchToRun(470) = 3996
    arrayCatchToRun(471) = 4016
    arrayCatchToRun(472) = 4024
    arrayCatchToRun(473) = 4038
    arrayCatchToRun(474) = 4041
    arrayCatchToRun(475) = 4045
    arrayCatchToRun(476) = 4061
    arrayCatchToRun(477) = 4068
    arrayCatchToRun(478) = 4075
    arrayCatchToRun(479) = 4122
    arrayCatchToRun(480) = 4129
    arrayCatchToRun(481) = 4145
    arrayCatchToRun(482) = 4188
    arrayCatchToRun(483) = 4192
    arrayCatchToRun(484) = 4195
    arrayCatchToRun(485) = 4204
    arrayCatchToRun(486) = 4211
    arrayCatchToRun(487) = 4219
    arrayCatchToRun(488) = 4228
    arrayCatchToRun(489) = 4231
    arrayCatchToRun(490) = 4244
    arrayCatchToRun(491) = 4257
    arrayCatchToRun(492) = 4262
    arrayCatchToRun(493) = 4286
    arrayCatchToRun(494) = 4327
    arrayCatchToRun(495) = 4334
    arrayCatchToRun(496) = 4350
    arrayCatchToRun(497) = 4454
    arrayCatchToRun(498) = 4456
    arrayCatchToRun(499) = 4471
    arrayCatchToRun(500) = 4481
    arrayCatchToRun(501) = 4550
    arrayCatchToRun(502) = 4556
    arrayCatchToRun(503) = 4606
    arrayCatchToRun(504) = 4607
    arrayCatchToRun(505) = 4611
    arrayCatchToRun(506) = 4626
    arrayCatchToRun(507) = 4635
    arrayCatchToRun(508) = 4683
    arrayCatchToRun(509) = 4685
    arrayCatchToRun(510) = 4758
    arrayCatchToRun(511) = 4814
    arrayCatchToRun(512) = 4818
    arrayCatchToRun(513) = 4819
    arrayCatchToRun(514) = 4821
    arrayCatchToRun(515) = 4826
    arrayCatchToRun(516) = 4846
    arrayCatchToRun(517) = 4871
    arrayCatchToRun(518) = 4979
    arrayCatchToRun(519) = 5005
    arrayCatchToRun(520) = 5048
    arrayCatchToRun(521) = 5109
    arrayCatchToRun(522) = 5141
    arrayCatchToRun(523) = 5174
    arrayCatchToRun(524) = 5182
    arrayCatchToRun(525) = 5188
    arrayCatchToRun(526) = 5194
    arrayCatchToRun(527) = 5234
    arrayCatchToRun(528) = 5259
    arrayCatchToRun(529) = 5325
    arrayCatchToRun(530) = 5331
    arrayCatchToRun(531) = 5342
    arrayCatchToRun(532) = 5356
    arrayCatchToRun(533) = 5386
    arrayCatchToRun(534) = 5410
    arrayCatchToRun(535) = 5442
    arrayCatchToRun(536) = 5479
    arrayCatchToRun(537) = 5536
    arrayCatchToRun(538) = 5538
    arrayCatchToRun(539) = 5547
    arrayCatchToRun(540) = 5579
    arrayCatchToRun(541) = 5629
    arrayCatchToRun(542) = 5737
    arrayCatchToRun(543) = 5775
    arrayCatchToRun(544) = 5777
    arrayCatchToRun(545) = 5799
    arrayCatchToRun(546) = 5802
    arrayCatchToRun(547) = 5806
    arrayCatchToRun(548) = 5859
    arrayCatchToRun(549) = 5918
    arrayCatchToRun(550) = 5933
    arrayCatchToRun(551) = 5965
    arrayCatchToRun(552) = 5976
    arrayCatchToRun(553) = 6000
    arrayCatchToRun(554) = 6013
    arrayCatchToRun(555) = 6025
    arrayCatchToRun(556) = 6044
    arrayCatchToRun(557) = 6049
    arrayCatchToRun(558) = 6072
    arrayCatchToRun(559) = 6094
    arrayCatchToRun(560) = 6099
    arrayCatchToRun(561) = 6110
    arrayCatchToRun(562) = 6202
    arrayCatchToRun(563) = 6280
    arrayCatchToRun(564) = 6302
    arrayCatchToRun(565) = 6382
    arrayCatchToRun(566) = 6510
    arrayCatchToRun(567) = 6526
    arrayCatchToRun(568) = 6577
    arrayCatchToRun(569) = 6614
    arrayCatchToRun(570) = 6617
    arrayCatchToRun(571) = 6622
    arrayCatchToRun(572) = 6637
    arrayCatchToRun(573) = 6638
    arrayCatchToRun(574) = 6671
    arrayCatchToRun(575) = 6690
    arrayCatchToRun(576) = 6696
    arrayCatchToRun(577) = 6757
    arrayCatchToRun(578) = 6760
    arrayCatchToRun(579) = 6806
    arrayCatchToRun(580) = 6819
    arrayCatchToRun(581) = 6832
    arrayCatchToRun(582) = 6872
    arrayCatchToRun(583) = 6881
    arrayCatchToRun(584) = 6890
    arrayCatchToRun(585) = 6907
    arrayCatchToRun(586) = 6919
    arrayCatchToRun(587) = 6938
    arrayCatchToRun(588) = 6952
    arrayCatchToRun(589) = 6972
    arrayCatchToRun(590) = 6984
    arrayCatchToRun(591) = 6986
    arrayCatchToRun(592) = 6994
    arrayCatchToRun(593) = 7005
    arrayCatchToRun(594) = 7016
    arrayCatchToRun(595) = 7029
    arrayCatchToRun(596) = 7061
    arrayCatchToRun(597) = 7070
    arrayCatchToRun(598) = 7077
    arrayCatchToRun(599) = 7078
    arrayCatchToRun(600) = 7088
    arrayCatchToRun(601) = 7096
    arrayCatchToRun(602) = 7097
    arrayCatchToRun(603) = 7102
    arrayCatchToRun(604) = 7144
    arrayCatchToRun(605) = 7145
    arrayCatchToRun(606) = 7146
    arrayCatchToRun(607) = 7156
    arrayCatchToRun(608) = 7160
    arrayCatchToRun(609) = 7166
    arrayCatchToRun(610) = 7175
    arrayCatchToRun(611) = 7209
    arrayCatchToRun(612) = 7212
    arrayCatchToRun(613) = 7219
    arrayCatchToRun(614) = 7238
    arrayCatchToRun(615) = 7245
    arrayCatchToRun(616) = 7267
    arrayCatchToRun(617) = 7279
    arrayCatchToRun(618) = 7298
    arrayCatchToRun(619) = 7306
    arrayCatchToRun(620) = 7311
    arrayCatchToRun(621) = 7326
    arrayCatchToRun(622) = 7327
    arrayCatchToRun(623) = 7333
    arrayCatchToRun(624) = 7337
    arrayCatchToRun(625) = 7351
    arrayCatchToRun(626) = 7352
    arrayCatchToRun(627) = 7360
    arrayCatchToRun(628) = 7366
    arrayCatchToRun(629) = 7369
    arrayCatchToRun(630) = 7372
    arrayCatchToRun(631) = 7374
    arrayCatchToRun(632) = 7378
    arrayCatchToRun(633) = 7385
    arrayCatchToRun(634) = 7391
    arrayCatchToRun(635) = 7393
    arrayCatchToRun(636) = 7414
    arrayCatchToRun(637) = 7426
    arrayCatchToRun(638) = 7446
    arrayCatchToRun(639) = 7448
    arrayCatchToRun(640) = 7453
    arrayCatchToRun(641) = 7455
    arrayCatchToRun(642) = 7467
    arrayCatchToRun(643) = 7487
    arrayCatchToRun(644) = 7491
    arrayCatchToRun(645) = 7498
    arrayCatchToRun(646) = 7500
    arrayCatchToRun(647) = 7511
    arrayCatchToRun(648) = 7522
    arrayCatchToRun(649) = 7536
    arrayCatchToRun(650) = 7550
    arrayCatchToRun(651) = 7605
    arrayCatchToRun(652) = 7653
    arrayCatchToRun(653) = 7667
    arrayCatchToRun(654) = 7682
    arrayCatchToRun(655) = 7709
    arrayCatchToRun(656) = 7736
    arrayCatchToRun(657) = 7748
    arrayCatchToRun(658) = 7752
    arrayCatchToRun(659) = 7767
    arrayCatchToRun(660) = 7791
    arrayCatchToRun(661) = 7824
    arrayCatchToRun(662) = 7859
    arrayCatchToRun(663) = 7863
    arrayCatchToRun(664) = 7884
    arrayCatchToRun(665) = 7921
    arrayCatchToRun(666) = 7942
    arrayCatchToRun(667) = 7955
    arrayCatchToRun(668) = 7974
    arrayCatchToRun(669) = 8000
    arrayCatchToRun(670) = 8005
    arrayCatchToRun(671) = 8028
    arrayCatchToRun(672) = 8092
    arrayCatchToRun(673) = 8117
    arrayCatchToRun(674) = 8155
    arrayCatchToRun(675) = 8158
    arrayCatchToRun(676) = 8224
    arrayCatchToRun(677) = 8252
    arrayCatchToRun(678) = 8266
    arrayCatchToRun(679) = 8281
    arrayCatchToRun(680) = 8299
    arrayCatchToRun(681) = 8329
    arrayCatchToRun(682) = 8338
    arrayCatchToRun(683) = 8363
    arrayCatchToRun(684) = 8383
    arrayCatchToRun(685) = 8546
    arrayCatchToRun(686) = 8603
    arrayCatchToRun(687) = 8619
    arrayCatchToRun(688) = 8638
    arrayCatchToRun(689) = 8656
    arrayCatchToRun(690) = 8914
    arrayCatchToRun(691) = 8945
    arrayCatchToRun(692) = 8946
    arrayCatchToRun(693) = 9058
    arrayCatchToRun(694) = 9078
    arrayCatchToRun(695) = 9094
    arrayCatchToRun(696) = 9130
    arrayCatchToRun(697) = 9138
    arrayCatchToRun(698) = 9152
    arrayCatchToRun(699) = 9194
    arrayCatchToRun(700) = 9217
    arrayCatchToRun(701) = 9230
    arrayCatchToRun(702) = 9262
    arrayCatchToRun(703) = 9294
    arrayCatchToRun(704) = 9305
    arrayCatchToRun(705) = 9312
    arrayCatchToRun(706) = 9407
    arrayCatchToRun(707) = 9428
    arrayCatchToRun(708) = 9447
    arrayCatchToRun(709) = 9516
    arrayCatchToRun(710) = 9561
    arrayCatchToRun(711) = 9583
    arrayCatchToRun(712) = 9623
    arrayCatchToRun(713) = 9730
    arrayCatchToRun(714) = 9733
    arrayCatchToRun(715) = 9746
    arrayCatchToRun(716) = 9788
    arrayCatchToRun(717) = 9801
    arrayCatchToRun(718) = 9829
    arrayCatchToRun(719) = 9911
    arrayCatchToRun(720) = 9959
    arrayCatchToRun(721) = 9964
    arrayCatchToRun(722) = 10017
    arrayCatchToRun(723) = 10023
    arrayCatchToRun(724) = 10026
    arrayCatchToRun(725) = 10047
    arrayCatchToRun(726) = 10049
    arrayCatchToRun(727) = 10071
    arrayCatchToRun(728) = 10080
    arrayCatchToRun(729) = 10108
    arrayCatchToRun(730) = 10157
    arrayCatchToRun(731) = 10183
    arrayCatchToRun(732) = 10213
    arrayCatchToRun(733) = 10259
    arrayCatchToRun(734) = 10328
    arrayCatchToRun(735) = 10336
    arrayCatchToRun(736) = 10346
    arrayCatchToRun(737) = 10380
    arrayCatchToRun(738) = 10383
    arrayCatchToRun(739) = 10418
    arrayCatchToRun(740) = 10430
    arrayCatchToRun(741) = 10444
    arrayCatchToRun(742) = 10458
    arrayCatchToRun(743) = 10474
    arrayCatchToRun(744) = 10502
    arrayCatchToRun(745) = 10513
    arrayCatchToRun(746) = 10515
    arrayCatchToRun(747) = 10550
    arrayCatchToRun(748) = 10560
    arrayCatchToRun(749) = 10586
    arrayCatchToRun(750) = 10599
    arrayCatchToRun(751) = 10603
    arrayCatchToRun(752) = 10627
    arrayCatchToRun(753) = 10653
    arrayCatchToRun(754) = 10662
    arrayCatchToRun(755) = 10672
    arrayCatchToRun(756) = 10696
    arrayCatchToRun(757) = 10705
    arrayCatchToRun(758) = 10750
    arrayCatchToRun(759) = 10792
    arrayCatchToRun(760) = 10812
    arrayCatchToRun(761) = 10814
    arrayCatchToRun(762) = 10877
    arrayCatchToRun(763) = 10920
    arrayCatchToRun(764) = 10979
    arrayCatchToRun(765) = 11013
    arrayCatchToRun(766) = 11021
    arrayCatchToRun(767) = 11024
    arrayCatchToRun(768) = 11036
    arrayCatchToRun(769) = 11043
    arrayCatchToRun(770) = 11078
    arrayCatchToRun(771) = 11101
    arrayCatchToRun(772) = 11129
    arrayCatchToRun(773) = 11145
    arrayCatchToRun(774) = 11150
    arrayCatchToRun(775) = 11173
    arrayCatchToRun(776) = 11190
    arrayCatchToRun(777) = 11198
    arrayCatchToRun(778) = 11213
    arrayCatchToRun(779) = 11311
    arrayCatchToRun(780) = 11330
    arrayCatchToRun(781) = 11340
    arrayCatchToRun(782) = 11361
    arrayCatchToRun(783) = 11367
    arrayCatchToRun(784) = 11371
    arrayCatchToRun(785) = 11376
    arrayCatchToRun(786) = 11381
    arrayCatchToRun(787) = 11408
    arrayCatchToRun(788) = 11421
    arrayCatchToRun(789) = 11454
    arrayCatchToRun(790) = 11471
    arrayCatchToRun(791) = 11483
    arrayCatchToRun(792) = 11500
    arrayCatchToRun(793) = 11510
    arrayCatchToRun(794) = 11539
    arrayCatchToRun(795) = 11545
    arrayCatchToRun(796) = 11551
    arrayCatchToRun(797) = 11572
    arrayCatchToRun(798) = 11600
    arrayCatchToRun(799) = 11604
    arrayCatchToRun(800) = 11649
    arrayCatchToRun(801) = 11650
    arrayCatchToRun(802) = 11656
    arrayCatchToRun(803) = 11658
    arrayCatchToRun(804) = 11661
    arrayCatchToRun(805) = 11685
    arrayCatchToRun(806) = 11700
    arrayCatchToRun(807) = 11711
    arrayCatchToRun(808) = 11720
    arrayCatchToRun(809) = 11759
    arrayCatchToRun(810) = 11765
    arrayCatchToRun(811) = 11770
    arrayCatchToRun(812) = 11783
    arrayCatchToRun(813) = 11801
    arrayCatchToRun(814) = 11815
    arrayCatchToRun(815) = 11843
    arrayCatchToRun(816) = 11845
    arrayCatchToRun(817) = 11846
    arrayCatchToRun(818) = 11851
    arrayCatchToRun(819) = 11906
    arrayCatchToRun(820) = 11907
    arrayCatchToRun(821) = 11937
    arrayCatchToRun(822) = 11947
    arrayCatchToRun(823) = 11960
    arrayCatchToRun(824) = 11973
    arrayCatchToRun(825) = 11986
    arrayCatchToRun(826) = 11995
    arrayCatchToRun(827) = 11999
    arrayCatchToRun(828) = 12000
    arrayCatchToRun(829) = 12015
    arrayCatchToRun(830) = 12038
    arrayCatchToRun(831) = 12053
    arrayCatchToRun(832) = 12064
    arrayCatchToRun(833) = 12082
    arrayCatchToRun(834) = 12083
    arrayCatchToRun(835) = 12085
    arrayCatchToRun(836) = 12089
    arrayCatchToRun(837) = 12090
    arrayCatchToRun(838) = 12098
    arrayCatchToRun(839) = 12107
    arrayCatchToRun(840) = 12112
    arrayCatchToRun(841) = 12118
    arrayCatchToRun(842) = 12130
    arrayCatchToRun(843) = 12137
    arrayCatchToRun(844) = 12138
    arrayCatchToRun(845) = 12143
    arrayCatchToRun(846) = 12153
    arrayCatchToRun(847) = 12159
    arrayCatchToRun(848) = 12190
    arrayCatchToRun(849) = 12232
    arrayCatchToRun(850) = 12242
    arrayCatchToRun(851) = 12269
    arrayCatchToRun(852) = 12274
    arrayCatchToRun(853) = 12278
    arrayCatchToRun(854) = 12284
    arrayCatchToRun(855) = 12290
    arrayCatchToRun(856) = 12317
    arrayCatchToRun(857) = 12321
    arrayCatchToRun(858) = 12328
    arrayCatchToRun(859) = 12339
    arrayCatchToRun(860) = 12344
    arrayCatchToRun(861) = 12351
    arrayCatchToRun(862) = 12371
    arrayCatchToRun(863) = 12374
    arrayCatchToRun(864) = 12384
    arrayCatchToRun(865) = 12402
    arrayCatchToRun(866) = 12406
    arrayCatchToRun(867) = 12426
    arrayCatchToRun(868) = 12434
    arrayCatchToRun(869) = 12440
    arrayCatchToRun(870) = 12452
    arrayCatchToRun(871) = 12469
    arrayCatchToRun(872) = 12502
    arrayCatchToRun(873) = 12515
    arrayCatchToRun(874) = 12516
    arrayCatchToRun(875) = 12528
    arrayCatchToRun(876) = 12530
    arrayCatchToRun(877) = 12543
    arrayCatchToRun(878) = 12555
    arrayCatchToRun(879) = 12557
    arrayCatchToRun(880) = 12563
    arrayCatchToRun(881) = 12564
    arrayCatchToRun(882) = 12571
    arrayCatchToRun(883) = 12591
    arrayCatchToRun(884) = 12628
    arrayCatchToRun(885) = 12659
    arrayCatchToRun(886) = 12662
    arrayCatchToRun(887) = 12678
    arrayCatchToRun(888) = 12733
    arrayCatchToRun(889) = 12737
    arrayCatchToRun(890) = 12756
    arrayCatchToRun(891) = 12759
    arrayCatchToRun(892) = 12760
    arrayCatchToRun(893) = 12763
    arrayCatchToRun(894) = 12782
    arrayCatchToRun(895) = 12803
    arrayCatchToRun(896) = 12848
    arrayCatchToRun(897) = 12858
    arrayCatchToRun(898) = 12877
    arrayCatchToRun(899) = 12878
    arrayCatchToRun(900) = 12902
    arrayCatchToRun(901) = 12904
    arrayCatchToRun(902) = 12918
    arrayCatchToRun(903) = 12920
    arrayCatchToRun(904) = 12924
    arrayCatchToRun(905) = 12936
    arrayCatchToRun(906) = 12944
    arrayCatchToRun(907) = 12945
    arrayCatchToRun(908) = 12987
    arrayCatchToRun(909) = 12988
    arrayCatchToRun(910) = 12995
    arrayCatchToRun(911) = 13008
    arrayCatchToRun(912) = 13049
    arrayCatchToRun(913) = 13058
    arrayCatchToRun(914) = 13105
    arrayCatchToRun(915) = 13108
    arrayCatchToRun(916) = 13133
    arrayCatchToRun(917) = 13135
    arrayCatchToRun(918) = 13144
    arrayCatchToRun(919) = 13166
    arrayCatchToRun(920) = 13187
    arrayCatchToRun(921) = 13189
    arrayCatchToRun(922) = 13196
    arrayCatchToRun(923) = 13201
    arrayCatchToRun(924) = 13230
    arrayCatchToRun(925) = 13260
    arrayCatchToRun(926) = 13284
    arrayCatchToRun(927) = 13308
    arrayCatchToRun(928) = 13371
    arrayCatchToRun(929) = 13381
    arrayCatchToRun(930) = 13400
    arrayCatchToRun(931) = 13414
    arrayCatchToRun(932) = 13422
    arrayCatchToRun(933) = 13424
    arrayCatchToRun(934) = 13427
    arrayCatchToRun(935) = 13448
    arrayCatchToRun(936) = 13450
    arrayCatchToRun(937) = 13453
    arrayCatchToRun(938) = 13483
    arrayCatchToRun(939) = 13484
    arrayCatchToRun(940) = 13497
    arrayCatchToRun(941) = 13505
    arrayCatchToRun(942) = 13508
    arrayCatchToRun(943) = 13513
    arrayCatchToRun(944) = 13527
    arrayCatchToRun(945) = 13536
    arrayCatchToRun(946) = 13538
    arrayCatchToRun(947) = 13558
    arrayCatchToRun(948) = 13559
    arrayCatchToRun(949) = 13565
    arrayCatchToRun(950) = 13566
    arrayCatchToRun(951) = 13570
    arrayCatchToRun(952) = 13581
    arrayCatchToRun(953) = 13591
    arrayCatchToRun(954) = 13594
    arrayCatchToRun(955) = 13604
    arrayCatchToRun(956) = 13614
    arrayCatchToRun(957) = 13616
    arrayCatchToRun(958) = 13672
    arrayCatchToRun(959) = 13676
    arrayCatchToRun(960) = 13681
    arrayCatchToRun(961) = 13690
    arrayCatchToRun(962) = 13703
    arrayCatchToRun(963) = 13704
    arrayCatchToRun(964) = 13714
    arrayCatchToRun(965) = 13725
    arrayCatchToRun(966) = 13732
    arrayCatchToRun(967) = 13737
    arrayCatchToRun(968) = 13739
    arrayCatchToRun(969) = 13758
    arrayCatchToRun(970) = 13761
    arrayCatchToRun(971) = 13770
    arrayCatchToRun(972) = 13775
    arrayCatchToRun(973) = 13791
    arrayCatchToRun(974) = 13793
    arrayCatchToRun(975) = 13796
    arrayCatchToRun(976) = 13797
    arrayCatchToRun(977) = 13798
    arrayCatchToRun(978) = 13800
    arrayCatchToRun(979) = 13813
    arrayCatchToRun(980) = 13844
    arrayCatchToRun(981) = 13849
    arrayCatchToRun(982) = 13865
    arrayCatchToRun(983) = 13866
    arrayCatchToRun(984) = 13921
    arrayCatchToRun(985) = 13928
    arrayCatchToRun(986) = 13949
    arrayCatchToRun(987) = 14005
    arrayCatchToRun(988) = 14011
    arrayCatchToRun(989) = 14012
    arrayCatchToRun(990) = 14019
    arrayCatchToRun(991) = 14034
    arrayCatchToRun(992) = 14079
    arrayCatchToRun(993) = 14091
    arrayCatchToRun(994) = 14094
    arrayCatchToRun(995) = 14107
    arrayCatchToRun(996) = 14114
    arrayCatchToRun(997) = 14118
    arrayCatchToRun(998) = 14120
    arrayCatchToRun(999) = 14123
    arrayCatchToRun(1000) = 14125
    arrayCatchToRun(1001) = 14132
    arrayCatchToRun(1002) = 14147
    arrayCatchToRun(1003) = 14152
    arrayCatchToRun(1004) = 14157
    arrayCatchToRun(1005) = 14245
    arrayCatchToRun(1006) = 14253
    arrayCatchToRun(1007) = 14266
    arrayCatchToRun(1008) = 14268
    arrayCatchToRun(1009) = 14274
    arrayCatchToRun(1010) = 14276
    arrayCatchToRun(1011) = 14281
    arrayCatchToRun(1012) = 14284
    arrayCatchToRun(1013) = 14294
    arrayCatchToRun(1014) = 14298
    arrayCatchToRun(1015) = 14301
    arrayCatchToRun(1016) = 14302
    arrayCatchToRun(1017) = 14308
    arrayCatchToRun(1018) = 14309
    arrayCatchToRun(1019) = 14323
    arrayCatchToRun(1020) = 14331
    arrayCatchToRun(1021) = 14336
    arrayCatchToRun(1022) = 14346
    arrayCatchToRun(1023) = 14352
    arrayCatchToRun(1024) = 14358
    arrayCatchToRun(1025) = 14363
    arrayCatchToRun(1026) = 14365
    arrayCatchToRun(1027) = 14375
    arrayCatchToRun(1028) = 14377
    arrayCatchToRun(1029) = 14380
    arrayCatchToRun(1030) = 14393
    arrayCatchToRun(1031) = 14403
    arrayCatchToRun(1032) = 14408
    arrayCatchToRun(1033) = 14413
    arrayCatchToRun(1034) = 14415
    arrayCatchToRun(1035) = 14421
    arrayCatchToRun(1036) = 14430
    arrayCatchToRun(1037) = 14468
    arrayCatchToRun(1038) = 14470
    arrayCatchToRun(1039) = 14488
    arrayCatchToRun(1040) = 14511
    arrayCatchToRun(1041) = 14524
    arrayCatchToRun(1042) = 14530
    arrayCatchToRun(1043) = 14539
    arrayCatchToRun(1044) = 14577
    arrayCatchToRun(1045) = 14582
    arrayCatchToRun(1046) = 14593
    arrayCatchToRun(1047) = 14600
    arrayCatchToRun(1048) = 14625
    arrayCatchToRun(1049) = 14649
    arrayCatchToRun(1050) = 14665
    arrayCatchToRun(1051) = 14681
    arrayCatchToRun(1052) = 14682
    arrayCatchToRun(1053) = 14709
    arrayCatchToRun(1054) = 14730
    arrayCatchToRun(1055) = 14752
    arrayCatchToRun(1056) = 14780
    arrayCatchToRun(1057) = 14814
    arrayCatchToRun(1058) = 14817
    arrayCatchToRun(1059) = 14828
    arrayCatchToRun(1060) = 14834
    arrayCatchToRun(1061) = 14846
    arrayCatchToRun(1062) = 14897
    arrayCatchToRun(1063) = 14899
    arrayCatchToRun(1064) = 14920
    arrayCatchToRun(1065) = 14927
    arrayCatchToRun(1066) = 14931
    arrayCatchToRun(1067) = 14946
    arrayCatchToRun(1068) = 14957
    arrayCatchToRun(1069) = 15016
    arrayCatchToRun(1070) = 15019
    arrayCatchToRun(1071) = 15053
    arrayCatchToRun(1072) = 15146
    arrayCatchToRun(1073) = 15159
    arrayCatchToRun(1074) = 15193
    arrayCatchToRun(1075) = 15204
    arrayCatchToRun(1076) = 15226
    arrayCatchToRun(1077) = 15337
    arrayCatchToRun(1078) = 15345
    arrayCatchToRun(1079) = 15374
    arrayCatchToRun(1080) = 15377
    arrayCatchToRun(1081) = 15409
    arrayCatchToRun(1082) = 15446
    arrayCatchToRun(1083) = 15449
    arrayCatchToRun(1084) = 15465
    arrayCatchToRun(1085) = 15485
    arrayCatchToRun(1086) = 15526
    arrayCatchToRun(1087) = 15585
    arrayCatchToRun(1088) = 15668
    arrayCatchToRun(1089) = 15676
    arrayCatchToRun(1090) = 15712
    arrayCatchToRun(1091) = 15721
    arrayCatchToRun(1092) = 15722
    arrayCatchToRun(1093) = 15730
    arrayCatchToRun(1094) = 15749
    arrayCatchToRun(1095) = 15750
    arrayCatchToRun(1096) = 15757
    arrayCatchToRun(1097) = 15760
    arrayCatchToRun(1098) = 15768
    arrayCatchToRun(1099) = 15775
    arrayCatchToRun(1100) = 15778
    arrayCatchToRun(1101) = 15781
    arrayCatchToRun(1102) = 15805
    arrayCatchToRun(1103) = 15821
    arrayCatchToRun(1104) = 15834
    arrayCatchToRun(1105) = 15835
    arrayCatchToRun(1106) = 15838
    arrayCatchToRun(1107) = 15853
    arrayCatchToRun(1108) = 15880
    arrayCatchToRun(1109) = 15902
    arrayCatchToRun(1110) = 15904
    arrayCatchToRun(1111) = 15916
    arrayCatchToRun(1112) = 15950
    arrayCatchToRun(1113) = 15970
    arrayCatchToRun(1114) = 15978
    arrayCatchToRun(1115) = 15995
    arrayCatchToRun(1116) = 16006
    arrayCatchToRun(1117) = 16025
    arrayCatchToRun(1118) = 16039
    arrayCatchToRun(1119) = 16054
    arrayCatchToRun(1120) = 16060
    arrayCatchToRun(1121) = 16093
    arrayCatchToRun(1122) = 16108
    arrayCatchToRun(1123) = 16116
    arrayCatchToRun(1124) = 16142
    arrayCatchToRun(1125) = 16187
    arrayCatchToRun(1126) = 16195
    arrayCatchToRun(1127) = 16211
    arrayCatchToRun(1128) = 16218
    arrayCatchToRun(1129) = 16224
    arrayCatchToRun(1130) = 16225
    arrayCatchToRun(1131) = 16232
    arrayCatchToRun(1132) = 16238
    arrayCatchToRun(1133) = 16247
    arrayCatchToRun(1134) = 16248
    arrayCatchToRun(1135) = 16259
    arrayCatchToRun(1136) = 16267
    arrayCatchToRun(1137) = 16274
    arrayCatchToRun(1138) = 16296
    arrayCatchToRun(1139) = 16302
    arrayCatchToRun(1140) = 16315
    arrayCatchToRun(1141) = 16320
    arrayCatchToRun(1142) = 16334
    arrayCatchToRun(1143) = 16337
    arrayCatchToRun(1144) = 16345
    arrayCatchToRun(1145) = 16348
    arrayCatchToRun(1146) = 16357
    arrayCatchToRun(1147) = 16359
    arrayCatchToRun(1148) = 16362
    arrayCatchToRun(1149) = 16380
    arrayCatchToRun(1150) = 16390
    arrayCatchToRun(1151) = 16393
    arrayCatchToRun(1152) = 16397
    arrayCatchToRun(1153) = 16401
    arrayCatchToRun(1154) = 16412
    arrayCatchToRun(1155) = 16415
    arrayCatchToRun(1156) = 16417
    arrayCatchToRun(1157) = 16425
    arrayCatchToRun(1158) = 16432
    arrayCatchToRun(1159) = 16439
    arrayCatchToRun(1160) = 16444
    arrayCatchToRun(1161) = 16449
    arrayCatchToRun(1162) = 16450
    arrayCatchToRun(1163) = 16451
    arrayCatchToRun(1164) = 16456
    arrayCatchToRun(1165) = 16463
    arrayCatchToRun(1166) = 16464
    arrayCatchToRun(1167) = 16476
    arrayCatchToRun(1168) = 16482
    arrayCatchToRun(1169) = 16485
    arrayCatchToRun(1170) = 16500
    arrayCatchToRun(1171) = 16513
    arrayCatchToRun(1172) = 16537
    arrayCatchToRun(1173) = 16538
    arrayCatchToRun(1174) = 16544
    arrayCatchToRun(1175) = 16550
    arrayCatchToRun(1176) = 16560
    arrayCatchToRun(1177) = 16562
    arrayCatchToRun(1178) = 16571
    arrayCatchToRun(1179) = 16577
    arrayCatchToRun(1180) = 16581
    arrayCatchToRun(1181) = 16596
    arrayCatchToRun(1182) = 16605
    arrayCatchToRun(1183) = 16606
    arrayCatchToRun(1184) = 16607
    arrayCatchToRun(1185) = 16609
    arrayCatchToRun(1186) = 16615
    arrayCatchToRun(1187) = 16617
    arrayCatchToRun(1188) = 16626
    arrayCatchToRun(1189) = 16628
    arrayCatchToRun(1190) = 16631
    arrayCatchToRun(1191) = 16634
    arrayCatchToRun(1192) = 16641
    arrayCatchToRun(1193) = 16665
    arrayCatchToRun(1194) = 16669
    arrayCatchToRun(1195) = 16688
    arrayCatchToRun(1196) = 16727
    arrayCatchToRun(1197) = 16728
    arrayCatchToRun(1198) = 16732
    arrayCatchToRun(1199) = 16741
    arrayCatchToRun(1200) = 16746
    arrayCatchToRun(1201) = 16757
    arrayCatchToRun(1202) = 16785
    arrayCatchToRun(1203) = 16791
    arrayCatchToRun(1204) = 16808
    arrayCatchToRun(1205) = 16811
    arrayCatchToRun(1206) = 16812
    arrayCatchToRun(1207) = 16818
    arrayCatchToRun(1208) = 16848
    arrayCatchToRun(1209) = 16850
    arrayCatchToRun(1210) = 16869
    arrayCatchToRun(1211) = 16875
    arrayCatchToRun(1212) = 16878
    arrayCatchToRun(1213) = 16895
    arrayCatchToRun(1214) = 16901
    arrayCatchToRun(1215) = 16904
    arrayCatchToRun(1216) = 16908
    arrayCatchToRun(1217) = 16925
    arrayCatchToRun(1218) = 16981
    arrayCatchToRun(1219) = 17000
    arrayCatchToRun(1220) = 17013
    arrayCatchToRun(1221) = 17019
    arrayCatchToRun(1222) = 17036
    arrayCatchToRun(1223) = 17050
    arrayCatchToRun(1224) = 17058
    arrayCatchToRun(1225) = 17077
    arrayCatchToRun(1226) = 17087
    arrayCatchToRun(1227) = 17138
    arrayCatchToRun(1228) = 17168
    arrayCatchToRun(1229) = 17189
    arrayCatchToRun(1230) = 17251
    arrayCatchToRun(1231) = 17255
    arrayCatchToRun(1232) = 17267
    arrayCatchToRun(1233) = 17287
    arrayCatchToRun(1234) = 17290
    arrayCatchToRun(1235) = 17293
    arrayCatchToRun(1236) = 17322
    arrayCatchToRun(1237) = 17332
    arrayCatchToRun(1238) = 17366
    arrayCatchToRun(1239) = 17375
    arrayCatchToRun(1240) = 17405
    arrayCatchToRun(1241) = 17442
    arrayCatchToRun(1242) = 17461
    arrayCatchToRun(1243) = 17498
    arrayCatchToRun(1244) = 17500
    arrayCatchToRun(1245) = 17514
    arrayCatchToRun(1246) = 17524
    arrayCatchToRun(1247) = 17552
    arrayCatchToRun(1248) = 17563
    arrayCatchToRun(1249) = 17565
    arrayCatchToRun(1250) = 17608
    arrayCatchToRun(1251) = 17617
    arrayCatchToRun(1252) = 17627
    arrayCatchToRun(1253) = 17665
    arrayCatchToRun(1254) = 17667
    arrayCatchToRun(1255) = 17669
    arrayCatchToRun(1256) = 17673
    arrayCatchToRun(1257) = 17682
    arrayCatchToRun(1258) = 17704
    arrayCatchToRun(1259) = 17713
    arrayCatchToRun(1260) = 17716
    arrayCatchToRun(1261) = 17719
    arrayCatchToRun(1262) = 17724
    arrayCatchToRun(1263) = 17727
    arrayCatchToRun(1264) = 17730
    arrayCatchToRun(1265) = 17751
    arrayCatchToRun(1266) = 17757
    arrayCatchToRun(1267) = 17758
    arrayCatchToRun(1268) = 17765
    arrayCatchToRun(1269) = 17767
    arrayCatchToRun(1270) = 17769
    arrayCatchToRun(1271) = 17783
    arrayCatchToRun(1272) = 17796
    arrayCatchToRun(1273) = 17804
    arrayCatchToRun(1274) = 17831
    arrayCatchToRun(1275) = 17842
    arrayCatchToRun(1276) = 17843
    arrayCatchToRun(1277) = 17845
    arrayCatchToRun(1278) = 17866
    arrayCatchToRun(1279) = 17913
    arrayCatchToRun(1280) = 17934
    arrayCatchToRun(1281) = 17941
    arrayCatchToRun(1282) = 17960
    arrayCatchToRun(1283) = 17973
    arrayCatchToRun(1284) = 17994
    arrayCatchToRun(1285) = 18010
    arrayCatchToRun(1286) = 18018
    arrayCatchToRun(1287) = 18034
    arrayCatchToRun(1288) = 18037
    arrayCatchToRun(1289) = 18038
    arrayCatchToRun(1290) = 18045
    arrayCatchToRun(1291) = 18053
    arrayCatchToRun(1292) = 18070
    arrayCatchToRun(1293) = 18093
    arrayCatchToRun(1294) = 18096
    arrayCatchToRun(1295) = 18122
    arrayCatchToRun(1296) = 18144
    arrayCatchToRun(1297) = 18166
    arrayCatchToRun(1298) = 18209
    arrayCatchToRun(1299) = 18226
    arrayCatchToRun(1300) = 18230
    arrayCatchToRun(1301) = 18245
    arrayCatchToRun(1302) = 18254
    arrayCatchToRun(1303) = 18330
    arrayCatchToRun(1304) = 18332
    arrayCatchToRun(1305) = 18340
    arrayCatchToRun(1306) = 18342
    arrayCatchToRun(1307) = 18350
    arrayCatchToRun(1308) = 18354
    arrayCatchToRun(1309) = 18356
    arrayCatchToRun(1310) = 18362
    arrayCatchToRun(1311) = 18405
    arrayCatchToRun(1312) = 18420
    arrayCatchToRun(1313) = 18504
    arrayCatchToRun(1314) = 18507
    arrayCatchToRun(1315) = 18509
    arrayCatchToRun(1316) = 18519
    arrayCatchToRun(1317) = 18521
    arrayCatchToRun(1318) = 18544
    arrayCatchToRun(1319) = 18551
    arrayCatchToRun(1320) = 18555
    arrayCatchToRun(1321) = 18560
    arrayCatchToRun(1322) = 18564
    arrayCatchToRun(1323) = 18592
    arrayCatchToRun(1324) = 18609
    arrayCatchToRun(1325) = 18615
    arrayCatchToRun(1326) = 18618
    arrayCatchToRun(1327) = 18624
    arrayCatchToRun(1328) = 18635
    arrayCatchToRun(1329) = 18658
    arrayCatchToRun(1330) = 18667
    arrayCatchToRun(1331) = 18684
    arrayCatchToRun(1332) = 18696
    arrayCatchToRun(1333) = 18702
    arrayCatchToRun(1334) = 18707
    arrayCatchToRun(1335) = 18709
    arrayCatchToRun(1336) = 18743
    arrayCatchToRun(1337) = 18751
    arrayCatchToRun(1338) = 18765
    arrayCatchToRun(1339) = 18788
    arrayCatchToRun(1340) = 18801
    arrayCatchToRun(1341) = 18802
    arrayCatchToRun(1342) = 18824
    arrayCatchToRun(1343) = 18832
    arrayCatchToRun(1344) = 18836
    arrayCatchToRun(1345) = 18837
    arrayCatchToRun(1346) = 18876
    arrayCatchToRun(1347) = 18898
    arrayCatchToRun(1348) = 18904
    arrayCatchToRun(1349) = 18905
    arrayCatchToRun(1350) = 18908
    arrayCatchToRun(1351) = 18911
    arrayCatchToRun(1352) = 18936
    arrayCatchToRun(1353) = 18937
    arrayCatchToRun(1354) = 18952
    arrayCatchToRun(1355) = 18955
    arrayCatchToRun(1356) = 18982
    arrayCatchToRun(1357) = 18985
    arrayCatchToRun(1358) = 18990
    arrayCatchToRun(1359) = 18993
    arrayCatchToRun(1360) = 18994
    arrayCatchToRun(1361) = 19009
    arrayCatchToRun(1362) = 19031
    arrayCatchToRun(1363) = 19050
    arrayCatchToRun(1364) = 19054
    arrayCatchToRun(1365) = 19063
    arrayCatchToRun(1366) = 19069
    arrayCatchToRun(1367) = 19077
    arrayCatchToRun(1368) = 19083
    arrayCatchToRun(1369) = 19090
    arrayCatchToRun(1370) = 19091
    arrayCatchToRun(1371) = 19096
    arrayCatchToRun(1372) = 19104
    arrayCatchToRun(1373) = 19126
    arrayCatchToRun(1374) = 19130
    arrayCatchToRun(1375) = 19133
    arrayCatchToRun(1376) = 19141
    arrayCatchToRun(1377) = 19147
    arrayCatchToRun(1378) = 19149
    arrayCatchToRun(1379) = 19173
    arrayCatchToRun(1380) = 19176
    arrayCatchToRun(1381) = 19197
    arrayCatchToRun(1382) = 19203
    arrayCatchToRun(1383) = 19221
    arrayCatchToRun(1384) = 19224
    arrayCatchToRun(1385) = 19262
    arrayCatchToRun(1386) = 19266
    arrayCatchToRun(1387) = 19268
    arrayCatchToRun(1388) = 19278
    arrayCatchToRun(1389) = 19294
    arrayCatchToRun(1390) = 19308
    arrayCatchToRun(1391) = 19320
    arrayCatchToRun(1392) = 19332
    arrayCatchToRun(1393) = 19338
    arrayCatchToRun(1394) = 19347
    arrayCatchToRun(1395) = 19396
    arrayCatchToRun(1396) = 19416
    arrayCatchToRun(1397) = 19420
    arrayCatchToRun(1398) = 19422
    arrayCatchToRun(1399) = 19429
    arrayCatchToRun(1400) = 19431
    arrayCatchToRun(1401) = 19442
    arrayCatchToRun(1402) = 19459
    arrayCatchToRun(1403) = 19474
    arrayCatchToRun(1404) = 19498
    arrayCatchToRun(1405) = 19500
    arrayCatchToRun(1406) = 19509
    arrayCatchToRun(1407) = 19511
    arrayCatchToRun(1408) = 19521
    arrayCatchToRun(1409) = 19528
    arrayCatchToRun(1410) = 19532
    arrayCatchToRun(1411) = 19548
    arrayCatchToRun(1412) = 19554
    arrayCatchToRun(1413) = 19581
    arrayCatchToRun(1414) = 19585
    arrayCatchToRun(1415) = 19596
    arrayCatchToRun(1416) = 19636
    arrayCatchToRun(1417) = 19651
    arrayCatchToRun(1418) = 19658
    arrayCatchToRun(1419) = 19668
    arrayCatchToRun(1420) = 19680
    arrayCatchToRun(1421) = 19683
    arrayCatchToRun(1422) = 19706
    arrayCatchToRun(1423) = 19717
    arrayCatchToRun(1424) = 19752
    arrayCatchToRun(1425) = 19762
    arrayCatchToRun(1426) = 19775
    arrayCatchToRun(1427) = 19782
    arrayCatchToRun(1428) = 19832
    arrayCatchToRun(1429) = 19838
    arrayCatchToRun(1430) = 19839
    arrayCatchToRun(1431) = 19853
    arrayCatchToRun(1432) = 19873
    arrayCatchToRun(1433) = 19906
    arrayCatchToRun(1434) = 19922
    arrayCatchToRun(1435) = 19930
    arrayCatchToRun(1436) = 19944
    arrayCatchToRun(1437) = 20010
    arrayCatchToRun(1438) = 20027
    arrayCatchToRun(1439) = 20072
    arrayCatchToRun(1440) = 20085
    arrayCatchToRun(1441) = 20099
    arrayCatchToRun(1442) = 20101
    arrayCatchToRun(1443) = 20108
    arrayCatchToRun(1444) = 20130
    arrayCatchToRun(1445) = 20154
    arrayCatchToRun(1446) = 20181
    arrayCatchToRun(1447) = 20187
    arrayCatchToRun(1448) = 20197
    arrayCatchToRun(1449) = 20206
    arrayCatchToRun(1450) = 20208
    arrayCatchToRun(1451) = 20213
    arrayCatchToRun(1452) = 20238
    arrayCatchToRun(1453) = 20244
    arrayCatchToRun(1454) = 20264
    arrayCatchToRun(1455) = 20287
    arrayCatchToRun(1456) = 20291
    arrayCatchToRun(1457) = 20296
    arrayCatchToRun(1458) = 20301
    arrayCatchToRun(1459) = 20336
    arrayCatchToRun(1460) = 20339
    arrayCatchToRun(1461) = 20343
    arrayCatchToRun(1462) = 20355
    arrayCatchToRun(1463) = 20363
    arrayCatchToRun(1464) = 20368
    arrayCatchToRun(1465) = 20404
    arrayCatchToRun(1466) = 20406
    arrayCatchToRun(1467) = 20409
    arrayCatchToRun(1468) = 20420
    arrayCatchToRun(1469) = 20429
    arrayCatchToRun(1470) = 20432
    arrayCatchToRun(1471) = 20434
    arrayCatchToRun(1472) = 20442
    arrayCatchToRun(1473) = 20446
    arrayCatchToRun(1474) = 20458
    arrayCatchToRun(1475) = 20463
    arrayCatchToRun(1476) = 20487
    arrayCatchToRun(1477) = 20492
    arrayCatchToRun(1478) = 20501
    arrayCatchToRun(1479) = 20502
    arrayCatchToRun(1480) = 20504
    arrayCatchToRun(1481) = 20505
    arrayCatchToRun(1482) = 20512
    arrayCatchToRun(1483) = 20519
    arrayCatchToRun(1484) = 20528
    arrayCatchToRun(1485) = 20533
    arrayCatchToRun(1486) = 20542
    arrayCatchToRun(1487) = 20547
    arrayCatchToRun(1488) = 20571
    arrayCatchToRun(1489) = 20574
    arrayCatchToRun(1490) = 20582
    arrayCatchToRun(1491) = 20590
    arrayCatchToRun(1492) = 20604
    arrayCatchToRun(1493) = 20608
    arrayCatchToRun(1494) = 20618
    arrayCatchToRun(1495) = 20626
    arrayCatchToRun(1496) = 20627
    arrayCatchToRun(1497) = 20632
    arrayCatchToRun(1498) = 20652
    arrayCatchToRun(1499) = 20667
    arrayCatchToRun(1500) = 20678
    arrayCatchToRun(1501) = 20728
    arrayCatchToRun(1502) = 20733
    arrayCatchToRun(1503) = 20745
    arrayCatchToRun(1504) = 20748
    arrayCatchToRun(1505) = 20755
    arrayCatchToRun(1506) = 20765
    arrayCatchToRun(1507) = 20778
    arrayCatchToRun(1508) = 20824
    arrayCatchToRun(1509) = 20859
    arrayCatchToRun(1510) = 20864
    arrayCatchToRun(1511) = 20888
    arrayCatchToRun(1512) = 20893
    arrayCatchToRun(1513) = 20920
    arrayCatchToRun(1514) = 20938
    arrayCatchToRun(1515) = 20945
    arrayCatchToRun(1516) = 20953
    arrayCatchToRun(1517) = 20957
    arrayCatchToRun(1518) = 20971
    arrayCatchToRun(1519) = 20972
    arrayCatchToRun(1520) = 20979
    arrayCatchToRun(1521) = 20986
    arrayCatchToRun(1522) = 21009
    arrayCatchToRun(1523) = 21020
    arrayCatchToRun(1524) = 21028
    arrayCatchToRun(1525) = 21040
    arrayCatchToRun(1526) = 21046
    arrayCatchToRun(1527) = 21049
    arrayCatchToRun(1528) = 21064
    arrayCatchToRun(1529) = 21073
    arrayCatchToRun(1530) = 21103
    arrayCatchToRun(1531) = 21109
    arrayCatchToRun(1532) = 21125
    arrayCatchToRun(1533) = 21145
    arrayCatchToRun(1534) = 21147
    arrayCatchToRun(1535) = 21151
    arrayCatchToRun(1536) = 21171
    arrayCatchToRun(1537) = 21174
    arrayCatchToRun(1538) = 21182
    arrayCatchToRun(1539) = 21187
    arrayCatchToRun(1540) = 21192
    arrayCatchToRun(1541) = 21198
    arrayCatchToRun(1542) = 21253
    arrayCatchToRun(1543) = 21265
    arrayCatchToRun(1544) = 21270
    arrayCatchToRun(1545) = 21272
    arrayCatchToRun(1546) = 21279
    arrayCatchToRun(1547) = 21284
    arrayCatchToRun(1548) = 21318
    arrayCatchToRun(1549) = 21344
    arrayCatchToRun(1550) = 21350
    arrayCatchToRun(1551) = 21373
    arrayCatchToRun(1552) = 21387
    arrayCatchToRun(1553) = 21391
    arrayCatchToRun(1554) = 21395
    arrayCatchToRun(1555) = 21439
    arrayCatchToRun(1556) = 21458
    arrayCatchToRun(1557) = 21488
    arrayCatchToRun(1558) = 21494
    arrayCatchToRun(1559) = 21501
    arrayCatchToRun(1560) = 21533
    arrayCatchToRun(1561) = 21552
    arrayCatchToRun(1562) = 21610
    arrayCatchToRun(1563) = 21644
    arrayCatchToRun(1564) = 21646
    arrayCatchToRun(1565) = 21671
    arrayCatchToRun(1566) = 21686
    arrayCatchToRun(1567) = 21687
    arrayCatchToRun(1568) = 21723
    arrayCatchToRun(1569) = 21726
    arrayCatchToRun(1570) = 21731
    arrayCatchToRun(1571) = 21762
    arrayCatchToRun(1572) = 21788
    arrayCatchToRun(1573) = 21835
    arrayCatchToRun(1574) = 21837
    arrayCatchToRun(1575) = 21845
    arrayCatchToRun(1576) = 21882
    arrayCatchToRun(1577) = 21894
    arrayCatchToRun(1578) = 21947
    arrayCatchToRun(1579) = 21967
    arrayCatchToRun(1580) = 21972
    arrayCatchToRun(1581) = 21989
    arrayCatchToRun(1582) = 22026
    arrayCatchToRun(1583) = 22051
    arrayCatchToRun(1584) = 22065
    arrayCatchToRun(1585) = 22082
    arrayCatchToRun(1586) = 22102
    arrayCatchToRun(1587) = 22110
    arrayCatchToRun(1588) = 22163
    arrayCatchToRun(1589) = 22171
    arrayCatchToRun(1590) = 22181
    arrayCatchToRun(1591) = 22217
    arrayCatchToRun(1592) = 22239
    arrayCatchToRun(1593) = 22243
    arrayCatchToRun(1594) = 22246
    arrayCatchToRun(1595) = 22251
    arrayCatchToRun(1596) = 22261
    arrayCatchToRun(1597) = 22264
    arrayCatchToRun(1598) = 22320
    arrayCatchToRun(1599) = 22323
    arrayCatchToRun(1600) = 22328
    arrayCatchToRun(1601) = 22365
    arrayCatchToRun(1602) = 22370
    arrayCatchToRun(1603) = 22385
    arrayCatchToRun(1604) = 22395
    arrayCatchToRun(1605) = 22405
    arrayCatchToRun(1606) = 22421
    arrayCatchToRun(1607) = 22443
    arrayCatchToRun(1608) = 22448
    arrayCatchToRun(1609) = 22450
    arrayCatchToRun(1610) = 22479
    arrayCatchToRun(1611) = 22481
    arrayCatchToRun(1612) = 22484
    arrayCatchToRun(1613) = 22485
    arrayCatchToRun(1614) = 22489
    arrayCatchToRun(1615) = 22491
    arrayCatchToRun(1616) = 22493
    arrayCatchToRun(1617) = 22494
    arrayCatchToRun(1618) = 22508
    arrayCatchToRun(1619) = 22510
    arrayCatchToRun(1620) = 22514
    arrayCatchToRun(1621) = 22515
    arrayCatchToRun(1622) = 22521
    arrayCatchToRun(1623) = 22526
    arrayCatchToRun(1624) = 22535
    arrayCatchToRun(1625) = 22541
    arrayCatchToRun(1626) = 22557
    arrayCatchToRun(1627) = 22560
    arrayCatchToRun(1628) = 22564
    arrayCatchToRun(1629) = 22565
    arrayCatchToRun(1630) = 22568
    arrayCatchToRun(1631) = 22580
    arrayCatchToRun(1632) = 22581
    arrayCatchToRun(1633) = 22593
    arrayCatchToRun(1634) = 22629
    arrayCatchToRun(1635) = 22643
    arrayCatchToRun(1636) = 22665
    arrayCatchToRun(1637) = 22679
    arrayCatchToRun(1638) = 22683
    arrayCatchToRun(1639) = 22685
    arrayCatchToRun(1640) = 22702
    arrayCatchToRun(1641) = 22706
    arrayCatchToRun(1642) = 22740
    arrayCatchToRun(1643) = 22742
    arrayCatchToRun(1644) = 22746
    arrayCatchToRun(1645) = 22748
    arrayCatchToRun(1646) = 22755
    arrayCatchToRun(1647) = 22778
    arrayCatchToRun(1648) = 22792
    arrayCatchToRun(1649) = 22804
    arrayCatchToRun(1650) = 22811
    arrayCatchToRun(1651) = 22815
    arrayCatchToRun(1652) = 22819
    arrayCatchToRun(1653) = 22825
    arrayCatchToRun(1654) = 22829
    arrayCatchToRun(1655) = 22832
    arrayCatchToRun(1656) = 22834
    arrayCatchToRun(1657) = 22835
    arrayCatchToRun(1658) = 22863
    arrayCatchToRun(1659) = 22878
    arrayCatchToRun(1660) = 22893
    arrayCatchToRun(1661) = 22908
    arrayCatchToRun(1662) = 22916
    arrayCatchToRun(1663) = 22925
    arrayCatchToRun(1664) = 22937
    arrayCatchToRun(1665) = 22953
    arrayCatchToRun(1666) = 22961
    arrayCatchToRun(1667) = 22963
    arrayCatchToRun(1668) = 22964
    arrayCatchToRun(1669) = 22967
    arrayCatchToRun(1670) = 23007
    arrayCatchToRun(1671) = 23013
    arrayCatchToRun(1672) = 23021
    arrayCatchToRun(1673) = 23032
    arrayCatchToRun(1674) = 23037
    arrayCatchToRun(1675) = 23059
    arrayCatchToRun(1676) = 23076
    arrayCatchToRun(1677) = 23093
    arrayCatchToRun(1678) = 23127
    arrayCatchToRun(1679) = 23138
    arrayCatchToRun(1680) = 23171
    arrayCatchToRun(1681) = 23178
    arrayCatchToRun(1682) = 23214
    arrayCatchToRun(1683) = 23275
    arrayCatchToRun(1684) = 23278
    arrayCatchToRun(1685) = 23301
    arrayCatchToRun(1686) = 23302
    arrayCatchToRun(1687) = 23317
    arrayCatchToRun(1688) = 23355
    arrayCatchToRun(1689) = 23359
    arrayCatchToRun(1690) = 23367
    arrayCatchToRun(1691) = 23370
    arrayCatchToRun(1692) = 23382
    arrayCatchToRun(1693) = 23392
    arrayCatchToRun(1694) = 23395
    arrayCatchToRun(1695) = 23409
    arrayCatchToRun(1696) = 23412
    arrayCatchToRun(1697) = 23430
    arrayCatchToRun(1698) = 23442
    arrayCatchToRun(1699) = 23449
    arrayCatchToRun(1700) = 23486
    arrayCatchToRun(1701) = 23489
    arrayCatchToRun(1702) = 23498
    arrayCatchToRun(1703) = 23504
    arrayCatchToRun(1704) = 23518
    arrayCatchToRun(1705) = 23521
    arrayCatchToRun(1706) = 23523
    arrayCatchToRun(1707) = 23536
    arrayCatchToRun(1708) = 23549
    arrayCatchToRun(1709) = 23550
    arrayCatchToRun(1710) = 23563
    arrayCatchToRun(1711) = 23577
    arrayCatchToRun(1712) = 23587
    arrayCatchToRun(1713) = 23598
    arrayCatchToRun(1714) = 23599
    arrayCatchToRun(1715) = 23602
    arrayCatchToRun(1716) = 23605
    arrayCatchToRun(1717) = 23608
    arrayCatchToRun(1718) = 23622
    arrayCatchToRun(1719) = 23623
    arrayCatchToRun(1720) = 23632
    arrayCatchToRun(1721) = 23633
    arrayCatchToRun(1722) = 23642
    arrayCatchToRun(1723) = 23651
    arrayCatchToRun(1724) = 23655
    arrayCatchToRun(1725) = 23659
    arrayCatchToRun(1726) = 23662
    arrayCatchToRun(1727) = 23670
    arrayCatchToRun(1728) = 23675
    arrayCatchToRun(1729) = 23695
    arrayCatchToRun(1730) = 23708
    arrayCatchToRun(1731) = 23713
    arrayCatchToRun(1732) = 23718
    arrayCatchToRun(1733) = 23729
    arrayCatchToRun(1734) = 23738
    arrayCatchToRun(1735) = 23745
    arrayCatchToRun(1736) = 23746
    arrayCatchToRun(1737) = 23748
    arrayCatchToRun(1738) = 23750
    arrayCatchToRun(1739) = 23751
    arrayCatchToRun(1740) = 23765
    arrayCatchToRun(1741) = 23788
    arrayCatchToRun(1742) = 23793
    arrayCatchToRun(1743) = 23795
    arrayCatchToRun(1744) = 23818
    arrayCatchToRun(1745) = 23823
    arrayCatchToRun(1746) = 23826
    arrayCatchToRun(1747) = 23837
    arrayCatchToRun(1748) = 23842
    arrayCatchToRun(1749) = 23843
    arrayCatchToRun(1750) = 23850
    arrayCatchToRun(1751) = 23874
    arrayCatchToRun(1752) = 23883
    arrayCatchToRun(1753) = 23884
    arrayCatchToRun(1754) = 23899
    arrayCatchToRun(1755) = 23901
    arrayCatchToRun(1756) = 23939
    arrayCatchToRun(1757) = 23947
    arrayCatchToRun(1758) = 23950
    arrayCatchToRun(1759) = 23954
    arrayCatchToRun(1760) = 23959
    arrayCatchToRun(1761) = 23964
    arrayCatchToRun(1762) = 23986
    arrayCatchToRun(1763) = 23992
    arrayCatchToRun(1764) = 24006
    arrayCatchToRun(1765) = 24012
    arrayCatchToRun(1766) = 24019
    arrayCatchToRun(1767) = 24036
    arrayCatchToRun(1768) = 24055
    arrayCatchToRun(1769) = 24063
    arrayCatchToRun(1770) = 24068
    arrayCatchToRun(1771) = 24078
    arrayCatchToRun(1772) = 24089
    arrayCatchToRun(1773) = 24095
    arrayCatchToRun(1774) = 24103
    arrayCatchToRun(1775) = 24116
    arrayCatchToRun(1776) = 24131
    arrayCatchToRun(1777) = 24134
    arrayCatchToRun(1778) = 24137
    arrayCatchToRun(1779) = 24147
    arrayCatchToRun(1780) = 24150
    arrayCatchToRun(1781) = 24153
    arrayCatchToRun(1782) = 24169
    arrayCatchToRun(1783) = 24170
    arrayCatchToRun(1784) = 24180
    arrayCatchToRun(1785) = 24183
    arrayCatchToRun(1786) = 24198
    arrayCatchToRun(1787) = 24201
    arrayCatchToRun(1788) = 24202
    arrayCatchToRun(1789) = 24203
    arrayCatchToRun(1790) = 24219
    arrayCatchToRun(1791) = 24231
    arrayCatchToRun(1792) = 24240
    arrayCatchToRun(1793) = 24247
    arrayCatchToRun(1794) = 24280
    arrayCatchToRun(1795) = 24293
    arrayCatchToRun(1796) = 24296
    arrayCatchToRun(1797) = 24303
    arrayCatchToRun(1798) = 24304
    arrayCatchToRun(1799) = 24307
    arrayCatchToRun(1800) = 24309
    arrayCatchToRun(1801) = 24312
    arrayCatchToRun(1802) = 24318
    arrayCatchToRun(1803) = 24327
    arrayCatchToRun(1804) = 24328
    arrayCatchToRun(1805) = 24338
    arrayCatchToRun(1806) = 24339
    arrayCatchToRun(1807) = 24346
    arrayCatchToRun(1808) = 24356
    arrayCatchToRun(1809) = 24360
    arrayCatchToRun(1810) = 24362
    arrayCatchToRun(1811) = 24363
    arrayCatchToRun(1812) = 24365
    arrayCatchToRun(1813) = 24373
    arrayCatchToRun(1814) = 24377
    arrayCatchToRun(1815) = 24380
    arrayCatchToRun(1816) = 24386
    arrayCatchToRun(1817) = 24397
    arrayCatchToRun(1818) = 24408
    arrayCatchToRun(1819) = 24422
    arrayCatchToRun(1820) = 24424
    arrayCatchToRun(1821) = 24431
    arrayCatchToRun(1822) = 24439
    arrayCatchToRun(1823) = 24455
    arrayCatchToRun(1824) = 24458
    arrayCatchToRun(1825) = 24460
    arrayCatchToRun(1826) = 24461
    arrayCatchToRun(1827) = 24467
    arrayCatchToRun(1828) = 24472
    arrayCatchToRun(1829) = 24473
    arrayCatchToRun(1830) = 24498
    arrayCatchToRun(1831) = 24500
    arrayCatchToRun(1832) = 24505
    arrayCatchToRun(1833) = 24516
    arrayCatchToRun(1834) = 24520
    arrayCatchToRun(1835) = 24532
    arrayCatchToRun(1836) = 24548
    arrayCatchToRun(1837) = 24553
    arrayCatchToRun(1838) = 24554
    arrayCatchToRun(1839) = 24556
    arrayCatchToRun(1840) = 24566
    arrayCatchToRun(1841) = 24576
    arrayCatchToRun(1842) = 24586
    arrayCatchToRun(1843) = 24588
    arrayCatchToRun(1844) = 24600
    arrayCatchToRun(1845) = 24603
    arrayCatchToRun(1846) = 24614
    arrayCatchToRun(1847) = 24616
    arrayCatchToRun(1848) = 24629
    arrayCatchToRun(1849) = 24633
    arrayCatchToRun(1850) = 24640
    arrayCatchToRun(1851) = 24646
    arrayCatchToRun(1852) = 24649
    arrayCatchToRun(1853) = 24650
    arrayCatchToRun(1854) = 24659
    arrayCatchToRun(1855) = 24661
    arrayCatchToRun(1856) = 24671
    arrayCatchToRun(1857) = 24674
    arrayCatchToRun(1858) = 24679
    arrayCatchToRun(1859) = 24685
    arrayCatchToRun(1860) = 24692
    arrayCatchToRun(1861) = 24716
    arrayCatchToRun(1862) = 24727
    arrayCatchToRun(1863) = 24742
    arrayCatchToRun(1864) = 24746
    arrayCatchToRun(1865) = 24748
    arrayCatchToRun(1866) = 24753
    arrayCatchToRun(1867) = 24760
    arrayCatchToRun(1868) = 24761
    arrayCatchToRun(1869) = 24764
    arrayCatchToRun(1870) = 24766
    arrayCatchToRun(1871) = 24771
    arrayCatchToRun(1872) = 24773
    arrayCatchToRun(1873) = 24791
    arrayCatchToRun(1874) = 24794
    arrayCatchToRun(1875) = 24799
    arrayCatchToRun(1876) = 24810
    arrayCatchToRun(1877) = 24824
    arrayCatchToRun(1878) = 24826
    arrayCatchToRun(1879) = 24836
    arrayCatchToRun(1880) = 24841
    arrayCatchToRun(1881) = 24850
    arrayCatchToRun(1882) = 24851
    arrayCatchToRun(1883) = 24852
    arrayCatchToRun(1884) = 24857
    arrayCatchToRun(1885) = 24858
    arrayCatchToRun(1886) = 24860
    arrayCatchToRun(1887) = 24865
    arrayCatchToRun(1888) = 24867
    arrayCatchToRun(1889) = 24868
    arrayCatchToRun(1890) = 24871
    arrayCatchToRun(1891) = 24876
    arrayCatchToRun(1892) = 24880
    arrayCatchToRun(1893) = 24881
    arrayCatchToRun(1894) = 24884
    arrayCatchToRun(1895) = 24885
    arrayCatchToRun(1896) = 24894
    arrayCatchToRun(1897) = 24896
    arrayCatchToRun(1898) = 24919
    arrayCatchToRun(1899) = 24920
    arrayCatchToRun(1900) = 24926
    arrayCatchToRun(1901) = 24928
    arrayCatchToRun(1902) = 24931
    arrayCatchToRun(1903) = 24933
    arrayCatchToRun(1904) = 24934
    arrayCatchToRun(1905) = 24938
    arrayCatchToRun(1906) = 24939
    arrayCatchToRun(1907) = 24940
    arrayCatchToRun(1908) = 24943
    arrayCatchToRun(1909) = 24944
    arrayCatchToRun(1910) = 24949
    arrayCatchToRun(1911) = 24950
    arrayCatchToRun(1912) = 24951
    arrayCatchToRun(1913) = 24953
    arrayCatchToRun(1914) = 24955
    arrayCatchToRun(1915) = 24957
    arrayCatchToRun(1916) = 24965
    arrayCatchToRun(1917) = 24966
    arrayCatchToRun(1918) = 24970
    arrayCatchToRun(1919) = 24973
    arrayCatchToRun(1920) = 24975
    arrayCatchToRun(1921) = 24988
    arrayCatchToRun(1922) = 24997
    arrayCatchToRun(1923) = 25007
    arrayCatchToRun(1924) = 25008
    arrayCatchToRun(1925) = 25014
    arrayCatchToRun(1926) = 25022
    arrayCatchToRun(1927) = 25034
    arrayCatchToRun(1928) = 25035
    arrayCatchToRun(1929) = 25038
    arrayCatchToRun(1930) = 25040
    arrayCatchToRun(1931) = 25041
    arrayCatchToRun(1932) = 25057
    arrayCatchToRun(1933) = 25066
    arrayCatchToRun(1934) = 25067
    arrayCatchToRun(1935) = 25070
    arrayCatchToRun(1936) = 25072
    arrayCatchToRun(1937) = 25077
    arrayCatchToRun(1938) = 25078
    arrayCatchToRun(1939) = 25081
    arrayCatchToRun(1940) = 25101
    arrayCatchToRun(1941) = 25106
    arrayCatchToRun(1942) = 25130
    arrayCatchToRun(1943) = 25133
    arrayCatchToRun(1944) = 25136
    arrayCatchToRun(1945) = 25146
    arrayCatchToRun(1946) = 25151
    arrayCatchToRun(1947) = 25160
    arrayCatchToRun(1948) = 25161
    arrayCatchToRun(1949) = 25163
    arrayCatchToRun(1950) = 25170
    arrayCatchToRun(1951) = 25171
End If

If intArrayToUse = 3 Then
    ReDim arrayCatchToRun(1594) As Integer
    arrayCatchToRun(0) = 25183
    arrayCatchToRun(1) = 25205
    arrayCatchToRun(2) = 25210
    arrayCatchToRun(3) = 25214
    arrayCatchToRun(4) = 25223
    arrayCatchToRun(5) = 25227
    arrayCatchToRun(6) = 25229
    arrayCatchToRun(7) = 25230
    arrayCatchToRun(8) = 25231
    arrayCatchToRun(9) = 25242
    arrayCatchToRun(10) = 25253
    arrayCatchToRun(11) = 25260
    arrayCatchToRun(12) = 25264
    arrayCatchToRun(13) = 25268
    arrayCatchToRun(14) = 25291
    arrayCatchToRun(15) = 25305
    arrayCatchToRun(16) = 25319
    arrayCatchToRun(17) = 25321
    arrayCatchToRun(18) = 25335
    arrayCatchToRun(19) = 25356
    arrayCatchToRun(20) = 25357
    arrayCatchToRun(21) = 25358
    arrayCatchToRun(22) = 25362
    arrayCatchToRun(23) = 25364
    arrayCatchToRun(24) = 25366
    arrayCatchToRun(25) = 25367
    arrayCatchToRun(26) = 25377
    arrayCatchToRun(27) = 25382
    arrayCatchToRun(28) = 25391
    arrayCatchToRun(29) = 25393
    arrayCatchToRun(30) = 25397
    arrayCatchToRun(31) = 25400
    arrayCatchToRun(32) = 25406
    arrayCatchToRun(33) = 25408
    arrayCatchToRun(34) = 25409
    arrayCatchToRun(35) = 25417
    arrayCatchToRun(36) = 25425
    arrayCatchToRun(37) = 25430
    arrayCatchToRun(38) = 25471
    arrayCatchToRun(39) = 25473
    arrayCatchToRun(40) = 25475
    arrayCatchToRun(41) = 25477
    arrayCatchToRun(42) = 25483
    arrayCatchToRun(43) = 25486
    arrayCatchToRun(44) = 25490
    arrayCatchToRun(45) = 25494
    arrayCatchToRun(46) = 25500
    arrayCatchToRun(47) = 25501
    arrayCatchToRun(48) = 25513
    arrayCatchToRun(49) = 25517
    arrayCatchToRun(50) = 25520
    arrayCatchToRun(51) = 25525
    arrayCatchToRun(52) = 25529
    arrayCatchToRun(53) = 25531
    arrayCatchToRun(54) = 25533
    arrayCatchToRun(55) = 25535
    arrayCatchToRun(56) = 25536
    arrayCatchToRun(57) = 25538
    arrayCatchToRun(58) = 25547
    arrayCatchToRun(59) = 25551
    arrayCatchToRun(60) = 25556
    arrayCatchToRun(61) = 25573
    arrayCatchToRun(62) = 25574
    arrayCatchToRun(63) = 25581
    arrayCatchToRun(64) = 25583
    arrayCatchToRun(65) = 25584
    arrayCatchToRun(66) = 25587
    arrayCatchToRun(67) = 25590
    arrayCatchToRun(68) = 25592
    arrayCatchToRun(69) = 25599
    arrayCatchToRun(70) = 25600
    arrayCatchToRun(71) = 25601
    arrayCatchToRun(72) = 25602
    arrayCatchToRun(73) = 25607
    arrayCatchToRun(74) = 25609
    arrayCatchToRun(75) = 25610
    arrayCatchToRun(76) = 25618
    arrayCatchToRun(77) = 25623
    arrayCatchToRun(78) = 25624
    arrayCatchToRun(79) = 25625
    arrayCatchToRun(80) = 25626
    arrayCatchToRun(81) = 25629
    arrayCatchToRun(82) = 25637
    arrayCatchToRun(83) = 25638
    arrayCatchToRun(84) = 25639
    arrayCatchToRun(85) = 25640
    arrayCatchToRun(86) = 25641
    arrayCatchToRun(87) = 25643
    arrayCatchToRun(88) = 25645
    arrayCatchToRun(89) = 25658
    arrayCatchToRun(90) = 25661
    arrayCatchToRun(91) = 25667
    arrayCatchToRun(92) = 25676
    arrayCatchToRun(93) = 25677
    arrayCatchToRun(94) = 25678
    arrayCatchToRun(95) = 25687
    arrayCatchToRun(96) = 25694
    arrayCatchToRun(97) = 25700
    arrayCatchToRun(98) = 25701
    arrayCatchToRun(99) = 25703
    arrayCatchToRun(100) = 25714
    arrayCatchToRun(101) = 25717
    arrayCatchToRun(102) = 25718
    arrayCatchToRun(103) = 25721
    arrayCatchToRun(104) = 25728
    arrayCatchToRun(105) = 25729
    arrayCatchToRun(106) = 25736
    arrayCatchToRun(107) = 25738
    arrayCatchToRun(108) = 25739
    arrayCatchToRun(109) = 25742
    arrayCatchToRun(110) = 25745
    arrayCatchToRun(111) = 25746
    arrayCatchToRun(112) = 25747
    arrayCatchToRun(113) = 25758
    arrayCatchToRun(114) = 25759
    arrayCatchToRun(115) = 25761
    arrayCatchToRun(116) = 25773
    arrayCatchToRun(117) = 25774
    arrayCatchToRun(118) = 25775
    arrayCatchToRun(119) = 25777
    arrayCatchToRun(120) = 25778
    arrayCatchToRun(121) = 25781
    arrayCatchToRun(122) = 25784
    arrayCatchToRun(123) = 25799
    arrayCatchToRun(124) = 25801
    arrayCatchToRun(125) = 25806
    arrayCatchToRun(126) = 25810
    arrayCatchToRun(127) = 25817
    arrayCatchToRun(128) = 25821
    arrayCatchToRun(129) = 25823
    arrayCatchToRun(130) = 25824
    arrayCatchToRun(131) = 25828
    arrayCatchToRun(132) = 25829
    arrayCatchToRun(133) = 25833
    arrayCatchToRun(134) = 25836
    arrayCatchToRun(135) = 25840
    arrayCatchToRun(136) = 25841
    arrayCatchToRun(137) = 25842
    arrayCatchToRun(138) = 25843
    arrayCatchToRun(139) = 25852
    arrayCatchToRun(140) = 25853
    arrayCatchToRun(141) = 25854
    arrayCatchToRun(142) = 25857
    arrayCatchToRun(143) = 25860
    arrayCatchToRun(144) = 25861
    arrayCatchToRun(145) = 25863
    arrayCatchToRun(146) = 25865
    arrayCatchToRun(147) = 25880
    arrayCatchToRun(148) = 25885
    arrayCatchToRun(149) = 25887
    arrayCatchToRun(150) = 25891
    arrayCatchToRun(151) = 25895
    arrayCatchToRun(152) = 25900
    arrayCatchToRun(153) = 25909
    arrayCatchToRun(154) = 25932
    arrayCatchToRun(155) = 25933
    arrayCatchToRun(156) = 25934
    arrayCatchToRun(157) = 25938
    arrayCatchToRun(158) = 25939
    arrayCatchToRun(159) = 25941
    arrayCatchToRun(160) = 25952
    arrayCatchToRun(161) = 25953
    arrayCatchToRun(162) = 25954
    arrayCatchToRun(163) = 25956
    arrayCatchToRun(164) = 25961
    arrayCatchToRun(165) = 25965
    arrayCatchToRun(166) = 25967
    arrayCatchToRun(167) = 25977
    arrayCatchToRun(168) = 25979
    arrayCatchToRun(169) = 25980
    arrayCatchToRun(170) = 25993
    arrayCatchToRun(171) = 25997
    arrayCatchToRun(172) = 26001
    arrayCatchToRun(173) = 26002
    arrayCatchToRun(174) = 26008
    arrayCatchToRun(175) = 26009
    arrayCatchToRun(176) = 26017
    arrayCatchToRun(177) = 26024
    arrayCatchToRun(178) = 26025
    arrayCatchToRun(179) = 26029
    arrayCatchToRun(180) = 26030
    arrayCatchToRun(181) = 26046
    arrayCatchToRun(182) = 26047
    arrayCatchToRun(183) = 26048
    arrayCatchToRun(184) = 26053
    arrayCatchToRun(185) = 26056
    arrayCatchToRun(186) = 26059
    arrayCatchToRun(187) = 26062
    arrayCatchToRun(188) = 26066
    arrayCatchToRun(189) = 26067
    arrayCatchToRun(190) = 26069
    arrayCatchToRun(191) = 26072
    arrayCatchToRun(192) = 26079
    arrayCatchToRun(193) = 26089
    arrayCatchToRun(194) = 26090
    arrayCatchToRun(195) = 26099
    arrayCatchToRun(196) = 26101
    arrayCatchToRun(197) = 26106
    arrayCatchToRun(198) = 26107
    arrayCatchToRun(199) = 26109
    arrayCatchToRun(200) = 26110
    arrayCatchToRun(201) = 26113
    arrayCatchToRun(202) = 26117
    arrayCatchToRun(203) = 26119
    arrayCatchToRun(204) = 26120
    arrayCatchToRun(205) = 26126
    arrayCatchToRun(206) = 26145
    arrayCatchToRun(207) = 26150
    arrayCatchToRun(208) = 26151
    arrayCatchToRun(209) = 26155
    arrayCatchToRun(210) = 26159
    arrayCatchToRun(211) = 26162
    arrayCatchToRun(212) = 26164
    arrayCatchToRun(213) = 26166
    arrayCatchToRun(214) = 26171
    arrayCatchToRun(215) = 26173
    arrayCatchToRun(216) = 26175
    arrayCatchToRun(217) = 26195
    arrayCatchToRun(218) = 26197
    arrayCatchToRun(219) = 26207
    arrayCatchToRun(220) = 26215
    arrayCatchToRun(221) = 26218
    arrayCatchToRun(222) = 26220
    arrayCatchToRun(223) = 26225
    arrayCatchToRun(224) = 26236
    arrayCatchToRun(225) = 26237
    arrayCatchToRun(226) = 26240
    arrayCatchToRun(227) = 26241
    arrayCatchToRun(228) = 26252
    arrayCatchToRun(229) = 26266
    arrayCatchToRun(230) = 26276
    arrayCatchToRun(231) = 26284
    arrayCatchToRun(232) = 26285
    arrayCatchToRun(233) = 26286
    arrayCatchToRun(234) = 26291
    arrayCatchToRun(235) = 26292
    arrayCatchToRun(236) = 26298
    arrayCatchToRun(237) = 26299
    arrayCatchToRun(238) = 26300
    arrayCatchToRun(239) = 26312
    arrayCatchToRun(240) = 26322
    arrayCatchToRun(241) = 26323
    arrayCatchToRun(242) = 26332
    arrayCatchToRun(243) = 26336
    arrayCatchToRun(244) = 26339
    arrayCatchToRun(245) = 26344
    arrayCatchToRun(246) = 26346
    arrayCatchToRun(247) = 26349
    arrayCatchToRun(248) = 26351
    arrayCatchToRun(249) = 26361
    arrayCatchToRun(250) = 26367
    arrayCatchToRun(251) = 26370
    arrayCatchToRun(252) = 26386
    arrayCatchToRun(253) = 26389
    arrayCatchToRun(254) = 26393
    arrayCatchToRun(255) = 26394
    arrayCatchToRun(256) = 26397
    arrayCatchToRun(257) = 26403
    arrayCatchToRun(258) = 26405
    arrayCatchToRun(259) = 26408
    arrayCatchToRun(260) = 26410
    arrayCatchToRun(261) = 26414
    arrayCatchToRun(262) = 26416
    arrayCatchToRun(263) = 26418
    arrayCatchToRun(264) = 26421
    arrayCatchToRun(265) = 26433
    arrayCatchToRun(266) = 26436
    arrayCatchToRun(267) = 26437
    arrayCatchToRun(268) = 26451
    arrayCatchToRun(269) = 26454
    arrayCatchToRun(270) = 26458
    arrayCatchToRun(271) = 26461
    arrayCatchToRun(272) = 26462
    arrayCatchToRun(273) = 26465
    arrayCatchToRun(274) = 26468
    arrayCatchToRun(275) = 26482
    arrayCatchToRun(276) = 26484
    arrayCatchToRun(277) = 26502
    arrayCatchToRun(278) = 26514
    arrayCatchToRun(279) = 26519
    arrayCatchToRun(280) = 26533
    arrayCatchToRun(281) = 26535
    arrayCatchToRun(282) = 26543
    arrayCatchToRun(283) = 26551
    arrayCatchToRun(284) = 26560
    arrayCatchToRun(285) = 26562
    arrayCatchToRun(286) = 26577
    arrayCatchToRun(287) = 26579
    arrayCatchToRun(288) = 26590
    arrayCatchToRun(289) = 26591
    arrayCatchToRun(290) = 26599
    arrayCatchToRun(291) = 26605
    arrayCatchToRun(292) = 26608
    arrayCatchToRun(293) = 26611
    arrayCatchToRun(294) = 26612
    arrayCatchToRun(295) = 26613
    arrayCatchToRun(296) = 26614
    arrayCatchToRun(297) = 26631
    arrayCatchToRun(298) = 26643
    arrayCatchToRun(299) = 26647
    arrayCatchToRun(300) = 26651
    arrayCatchToRun(301) = 26654
    arrayCatchToRun(302) = 26655
    arrayCatchToRun(303) = 26657
    arrayCatchToRun(304) = 26664
    arrayCatchToRun(305) = 26678
    arrayCatchToRun(306) = 26679
    arrayCatchToRun(307) = 26693
    arrayCatchToRun(308) = 26697
    arrayCatchToRun(309) = 26699
    arrayCatchToRun(310) = 26708
    arrayCatchToRun(311) = 26715
    arrayCatchToRun(312) = 26719
    arrayCatchToRun(313) = 26722
    arrayCatchToRun(314) = 26723
    arrayCatchToRun(315) = 26724
    arrayCatchToRun(316) = 26725
    arrayCatchToRun(317) = 26726
    arrayCatchToRun(318) = 26727
    arrayCatchToRun(319) = 26730
    arrayCatchToRun(320) = 26733
    arrayCatchToRun(321) = 26734
    arrayCatchToRun(322) = 26737
    arrayCatchToRun(323) = 26752
    arrayCatchToRun(324) = 26756
    arrayCatchToRun(325) = 26765
    arrayCatchToRun(326) = 26767
    arrayCatchToRun(327) = 26770
    arrayCatchToRun(328) = 26776
    arrayCatchToRun(329) = 26790
    arrayCatchToRun(330) = 26796
    arrayCatchToRun(331) = 26806
    arrayCatchToRun(332) = 26809
    arrayCatchToRun(333) = 26812
    arrayCatchToRun(334) = 26831
    arrayCatchToRun(335) = 26832
    arrayCatchToRun(336) = 26862
    arrayCatchToRun(337) = 26865
    arrayCatchToRun(338) = 26866
    arrayCatchToRun(339) = 26871
    arrayCatchToRun(340) = 26881
    arrayCatchToRun(341) = 26882
    arrayCatchToRun(342) = 26886
    arrayCatchToRun(343) = 26901
    arrayCatchToRun(344) = 26902
    arrayCatchToRun(345) = 26905
    arrayCatchToRun(346) = 26906
    arrayCatchToRun(347) = 26918
    arrayCatchToRun(348) = 26932
    arrayCatchToRun(349) = 26948
    arrayCatchToRun(350) = 26963
    arrayCatchToRun(351) = 26965
    arrayCatchToRun(352) = 26968
    arrayCatchToRun(353) = 26969
    arrayCatchToRun(354) = 26970
    arrayCatchToRun(355) = 26972
    arrayCatchToRun(356) = 26975
    arrayCatchToRun(357) = 26990
    arrayCatchToRun(358) = 26995
    arrayCatchToRun(359) = 26996
    arrayCatchToRun(360) = 27002
    arrayCatchToRun(361) = 27006
    arrayCatchToRun(362) = 27013
    arrayCatchToRun(363) = 27032
    arrayCatchToRun(364) = 27040
    arrayCatchToRun(365) = 27055
    arrayCatchToRun(366) = 27071
    arrayCatchToRun(367) = 27072
    arrayCatchToRun(368) = 27079
    arrayCatchToRun(369) = 27089
    arrayCatchToRun(370) = 27101
    arrayCatchToRun(371) = 27121
    arrayCatchToRun(372) = 27146
    arrayCatchToRun(373) = 27149
    arrayCatchToRun(374) = 27151
    arrayCatchToRun(375) = 27160
    arrayCatchToRun(376) = 27170
    arrayCatchToRun(377) = 27172
    arrayCatchToRun(378) = 27186
    arrayCatchToRun(379) = 27187
    arrayCatchToRun(380) = 27191
    arrayCatchToRun(381) = 27206
    arrayCatchToRun(382) = 27213
    arrayCatchToRun(383) = 27222
    arrayCatchToRun(384) = 27233
    arrayCatchToRun(385) = 27234
    arrayCatchToRun(386) = 27239
    arrayCatchToRun(387) = 27240
    arrayCatchToRun(388) = 27255
    arrayCatchToRun(389) = 27257
    arrayCatchToRun(390) = 27258
    arrayCatchToRun(391) = 27261
    arrayCatchToRun(392) = 27267
    arrayCatchToRun(393) = 27274
    arrayCatchToRun(394) = 27285
    arrayCatchToRun(395) = 27290
    arrayCatchToRun(396) = 27294
    arrayCatchToRun(397) = 27299
    arrayCatchToRun(398) = 27304
    arrayCatchToRun(399) = 27311
    arrayCatchToRun(400) = 27319
    arrayCatchToRun(401) = 27325
    arrayCatchToRun(402) = 27347
    arrayCatchToRun(403) = 27348
    arrayCatchToRun(404) = 27351
    arrayCatchToRun(405) = 27354
    arrayCatchToRun(406) = 27355
    arrayCatchToRun(407) = 27358
    arrayCatchToRun(408) = 27361
    arrayCatchToRun(409) = 27370
    arrayCatchToRun(410) = 27377
    arrayCatchToRun(411) = 27381
    arrayCatchToRun(412) = 27387
    arrayCatchToRun(413) = 27394
    arrayCatchToRun(414) = 27399
    arrayCatchToRun(415) = 27414
    arrayCatchToRun(416) = 27416
    arrayCatchToRun(417) = 27420
    arrayCatchToRun(418) = 27428
    arrayCatchToRun(419) = 27443
    arrayCatchToRun(420) = 27447
    arrayCatchToRun(421) = 27451
    arrayCatchToRun(422) = 27456
    arrayCatchToRun(423) = 27457
    arrayCatchToRun(424) = 27471
    arrayCatchToRun(425) = 27473
    arrayCatchToRun(426) = 27477
    arrayCatchToRun(427) = 27485
    arrayCatchToRun(428) = 27488
    arrayCatchToRun(429) = 27489
    arrayCatchToRun(430) = 27492
    arrayCatchToRun(431) = 27494
    arrayCatchToRun(432) = 27499
    arrayCatchToRun(433) = 27500
    arrayCatchToRun(434) = 27501
    arrayCatchToRun(435) = 27503
    arrayCatchToRun(436) = 27507
    arrayCatchToRun(437) = 27509
    arrayCatchToRun(438) = 27510
    arrayCatchToRun(439) = 27511
    arrayCatchToRun(440) = 27512
    arrayCatchToRun(441) = 27519
    arrayCatchToRun(442) = 27530
    arrayCatchToRun(443) = 27531
    arrayCatchToRun(444) = 27532
    arrayCatchToRun(445) = 27545
    arrayCatchToRun(446) = 27546
    arrayCatchToRun(447) = 27548
    arrayCatchToRun(448) = 27551
    arrayCatchToRun(449) = 27555
    arrayCatchToRun(450) = 27557
    arrayCatchToRun(451) = 27559
    arrayCatchToRun(452) = 27561
    arrayCatchToRun(453) = 27579
    arrayCatchToRun(454) = 27581
    arrayCatchToRun(455) = 27588
    arrayCatchToRun(456) = 27601
    arrayCatchToRun(457) = 27612
    arrayCatchToRun(458) = 27620
    arrayCatchToRun(459) = 27624
    arrayCatchToRun(460) = 27628
    arrayCatchToRun(461) = 27633
    arrayCatchToRun(462) = 27636
    arrayCatchToRun(463) = 27640
    arrayCatchToRun(464) = 27641
    arrayCatchToRun(465) = 27646
    arrayCatchToRun(466) = 27650
    arrayCatchToRun(467) = 27659
    arrayCatchToRun(468) = 27670
    arrayCatchToRun(469) = 27674
    arrayCatchToRun(470) = 27676
    arrayCatchToRun(471) = 27680
    arrayCatchToRun(472) = 27682
    arrayCatchToRun(473) = 27689
    arrayCatchToRun(474) = 27709
    arrayCatchToRun(475) = 27712
    arrayCatchToRun(476) = 27737
    arrayCatchToRun(477) = 27740
    arrayCatchToRun(478) = 27745
    arrayCatchToRun(479) = 27768
    arrayCatchToRun(480) = 27778
    arrayCatchToRun(481) = 27785
    arrayCatchToRun(482) = 27797
    arrayCatchToRun(483) = 27821
    arrayCatchToRun(484) = 27822
    arrayCatchToRun(485) = 27823
    arrayCatchToRun(486) = 27828
    arrayCatchToRun(487) = 27830
    arrayCatchToRun(488) = 27831
    arrayCatchToRun(489) = 27841
    arrayCatchToRun(490) = 27842
    arrayCatchToRun(491) = 27846
    arrayCatchToRun(492) = 27850
    arrayCatchToRun(493) = 27852
    arrayCatchToRun(494) = 27858
    arrayCatchToRun(495) = 27873
    arrayCatchToRun(496) = 27875
    arrayCatchToRun(497) = 27879
    arrayCatchToRun(498) = 27885
    arrayCatchToRun(499) = 27887
    arrayCatchToRun(500) = 27913
    arrayCatchToRun(501) = 27917
    arrayCatchToRun(502) = 27919
    arrayCatchToRun(503) = 27925
    arrayCatchToRun(504) = 27930
    arrayCatchToRun(505) = 27946
    arrayCatchToRun(506) = 27947
    arrayCatchToRun(507) = 27951
    arrayCatchToRun(508) = 27956
    arrayCatchToRun(509) = 27966
    arrayCatchToRun(510) = 27968
    arrayCatchToRun(511) = 27973
    arrayCatchToRun(512) = 27974
    arrayCatchToRun(513) = 27982
    arrayCatchToRun(514) = 27987
    arrayCatchToRun(515) = 27993
    arrayCatchToRun(516) = 28006
    arrayCatchToRun(517) = 28018
    arrayCatchToRun(518) = 28021
    arrayCatchToRun(519) = 28030
    arrayCatchToRun(520) = 28035
    arrayCatchToRun(521) = 28054
    arrayCatchToRun(522) = 28060
    arrayCatchToRun(523) = 28091
    arrayCatchToRun(524) = 28108
    arrayCatchToRun(525) = 28111
    arrayCatchToRun(526) = 28114
    arrayCatchToRun(527) = 28116
    arrayCatchToRun(528) = 28137
    arrayCatchToRun(529) = 28139
    arrayCatchToRun(530) = 28143
    arrayCatchToRun(531) = 28146
    arrayCatchToRun(532) = 28147
    arrayCatchToRun(533) = 28160
    arrayCatchToRun(534) = 28181
    arrayCatchToRun(535) = 28182
    arrayCatchToRun(536) = 28184
    arrayCatchToRun(537) = 28189
    arrayCatchToRun(538) = 28209
    arrayCatchToRun(539) = 28210
    arrayCatchToRun(540) = 28233
    arrayCatchToRun(541) = 28234
    arrayCatchToRun(542) = 28239
    arrayCatchToRun(543) = 28251
    arrayCatchToRun(544) = 28256
    arrayCatchToRun(545) = 28260
    arrayCatchToRun(546) = 28277
    arrayCatchToRun(547) = 28280
    arrayCatchToRun(548) = 28288
    arrayCatchToRun(549) = 28302
    arrayCatchToRun(550) = 28317
    arrayCatchToRun(551) = 28325
    arrayCatchToRun(552) = 28329
    arrayCatchToRun(553) = 28332
    arrayCatchToRun(554) = 28336
    arrayCatchToRun(555) = 28337
    arrayCatchToRun(556) = 28339
    arrayCatchToRun(557) = 28345
    arrayCatchToRun(558) = 28351
    arrayCatchToRun(559) = 28354
    arrayCatchToRun(560) = 28363
    arrayCatchToRun(561) = 28368
    arrayCatchToRun(562) = 28370
    arrayCatchToRun(563) = 28374
    arrayCatchToRun(564) = 28375
    arrayCatchToRun(565) = 28377
    arrayCatchToRun(566) = 28379
    arrayCatchToRun(567) = 28400
    arrayCatchToRun(568) = 28407
    arrayCatchToRun(569) = 28413
    arrayCatchToRun(570) = 28418
    arrayCatchToRun(571) = 28427
    arrayCatchToRun(572) = 28430
    arrayCatchToRun(573) = 28432
    arrayCatchToRun(574) = 28434
    arrayCatchToRun(575) = 28465
    arrayCatchToRun(576) = 28471
    arrayCatchToRun(577) = 28475
    arrayCatchToRun(578) = 28485
    arrayCatchToRun(579) = 28493
    arrayCatchToRun(580) = 28499
    arrayCatchToRun(581) = 28510
    arrayCatchToRun(582) = 28518
    arrayCatchToRun(583) = 28524
    arrayCatchToRun(584) = 28548
    arrayCatchToRun(585) = 28549
    arrayCatchToRun(586) = 28554
    arrayCatchToRun(587) = 28560
    arrayCatchToRun(588) = 28574
    arrayCatchToRun(589) = 28581
    arrayCatchToRun(590) = 28619
    arrayCatchToRun(591) = 28625
    arrayCatchToRun(592) = 28685
    arrayCatchToRun(593) = 28687
    arrayCatchToRun(594) = 28697
    arrayCatchToRun(595) = 28698
    arrayCatchToRun(596) = 28702
    arrayCatchToRun(597) = 28706
    arrayCatchToRun(598) = 28713
    arrayCatchToRun(599) = 28722
    arrayCatchToRun(600) = 28740
    arrayCatchToRun(601) = 28775
    arrayCatchToRun(602) = 28795
    arrayCatchToRun(603) = 9
    arrayCatchToRun(604) = 7
    arrayCatchToRun(605) = 110
    arrayCatchToRun(606) = 95
    arrayCatchToRun(607) = 78
    arrayCatchToRun(608) = 156
    arrayCatchToRun(609) = 173
    arrayCatchToRun(610) = 196
    arrayCatchToRun(611) = 190
    arrayCatchToRun(612) = 195
    arrayCatchToRun(613) = 217
    arrayCatchToRun(614) = 231
    arrayCatchToRun(615) = 266
    arrayCatchToRun(616) = 250
    arrayCatchToRun(617) = 289
    arrayCatchToRun(618) = 292
    arrayCatchToRun(619) = 274
    arrayCatchToRun(620) = 331
    arrayCatchToRun(621) = 363
    arrayCatchToRun(622) = 386
    arrayCatchToRun(623) = 458
    arrayCatchToRun(624) = 309
    arrayCatchToRun(625) = 395
    arrayCatchToRun(626) = 586
    arrayCatchToRun(627) = 612
    arrayCatchToRun(628) = 631
    arrayCatchToRun(629) = 687
    arrayCatchToRun(630) = 671
    arrayCatchToRun(631) = 692
    arrayCatchToRun(632) = 718
    arrayCatchToRun(633) = 728
    arrayCatchToRun(634) = 729
    arrayCatchToRun(635) = 754
    arrayCatchToRun(636) = 781
    arrayCatchToRun(637) = 772
    arrayCatchToRun(638) = 790
    arrayCatchToRun(639) = 802
    arrayCatchToRun(640) = 817
    arrayCatchToRun(641) = 849
    arrayCatchToRun(642) = 869
    arrayCatchToRun(643) = 882
    arrayCatchToRun(644) = 901
    arrayCatchToRun(645) = 946
    arrayCatchToRun(646) = 918
    arrayCatchToRun(647) = 938
    arrayCatchToRun(648) = 939
    arrayCatchToRun(649) = 950
    arrayCatchToRun(650) = 990
    arrayCatchToRun(651) = 977
    arrayCatchToRun(652) = 965
    arrayCatchToRun(653) = 998
    arrayCatchToRun(654) = 1044
    arrayCatchToRun(655) = 1030
    arrayCatchToRun(656) = 1208
    arrayCatchToRun(657) = 1057
    arrayCatchToRun(658) = 1088
    arrayCatchToRun(659) = 1129
    arrayCatchToRun(660) = 1115
    arrayCatchToRun(661) = 1081
    arrayCatchToRun(662) = 1137
    arrayCatchToRun(663) = 1149
    arrayCatchToRun(664) = 1293
    arrayCatchToRun(665) = 1155
    arrayCatchToRun(666) = 1165
    arrayCatchToRun(667) = 1150
    arrayCatchToRun(668) = 1224
    arrayCatchToRun(669) = 1202
    arrayCatchToRun(670) = 1251
    arrayCatchToRun(671) = 1234
    arrayCatchToRun(672) = 1247
    arrayCatchToRun(673) = 1242
    arrayCatchToRun(674) = 1260
    arrayCatchToRun(675) = 1271
    arrayCatchToRun(676) = 1294
    arrayCatchToRun(677) = 1319
    arrayCatchToRun(678) = 1328
    arrayCatchToRun(679) = 1338
    arrayCatchToRun(680) = 1346
    arrayCatchToRun(681) = 1374
    arrayCatchToRun(682) = 1381
    arrayCatchToRun(683) = 1393
    arrayCatchToRun(684) = 1423
    arrayCatchToRun(685) = 1416
    arrayCatchToRun(686) = 1444
    arrayCatchToRun(687) = 1411
    arrayCatchToRun(688) = 1450
    arrayCatchToRun(689) = 1445
    arrayCatchToRun(690) = 1456
    arrayCatchToRun(691) = 1570
    arrayCatchToRun(692) = 1614
    arrayCatchToRun(693) = 1615
    arrayCatchToRun(694) = 1619
    arrayCatchToRun(695) = 1631
    arrayCatchToRun(696) = 1663
    arrayCatchToRun(697) = 1694
    arrayCatchToRun(698) = 1690
    arrayCatchToRun(699) = 1809
    arrayCatchToRun(700) = 1711
    arrayCatchToRun(701) = 2065
    arrayCatchToRun(702) = 2078
    arrayCatchToRun(703) = 2096
    arrayCatchToRun(704) = 2138
    arrayCatchToRun(705) = 2144
    arrayCatchToRun(706) = 2150
    arrayCatchToRun(707) = 2163
    arrayCatchToRun(708) = 2170
    arrayCatchToRun(709) = 2182
    arrayCatchToRun(710) = 2237
    arrayCatchToRun(711) = 2259
    arrayCatchToRun(712) = 2262
    arrayCatchToRun(713) = 2306
    arrayCatchToRun(714) = 2536
    arrayCatchToRun(715) = 2565
    arrayCatchToRun(716) = 2335
    arrayCatchToRun(717) = 2438
    arrayCatchToRun(718) = 2382
    arrayCatchToRun(719) = 2420
    arrayCatchToRun(720) = 2212
    arrayCatchToRun(721) = 2444
    arrayCatchToRun(722) = 2488
    arrayCatchToRun(723) = 2712
    arrayCatchToRun(724) = 2469
    arrayCatchToRun(725) = 2603
    arrayCatchToRun(726) = 2487
    arrayCatchToRun(727) = 2533
    arrayCatchToRun(728) = 2494
    arrayCatchToRun(729) = 2583
    arrayCatchToRun(730) = 2490
    arrayCatchToRun(731) = 2623
    arrayCatchToRun(732) = 2402
    arrayCatchToRun(733) = 2464
    arrayCatchToRun(734) = 2501
    arrayCatchToRun(735) = 2735
    arrayCatchToRun(736) = 2633
    arrayCatchToRun(737) = 2358
    arrayCatchToRun(738) = 2723
    arrayCatchToRun(739) = 2652
    arrayCatchToRun(740) = 2828
    arrayCatchToRun(741) = 2931
    arrayCatchToRun(742) = 3264
    arrayCatchToRun(743) = 2937
    arrayCatchToRun(744) = 2833
    arrayCatchToRun(745) = 2933
    arrayCatchToRun(746) = 3092
    arrayCatchToRun(747) = 2861
    arrayCatchToRun(748) = 3004
    arrayCatchToRun(749) = 2831
    arrayCatchToRun(750) = 2944
    arrayCatchToRun(751) = 3188
    arrayCatchToRun(752) = 3181
    arrayCatchToRun(753) = 3170
    arrayCatchToRun(754) = 3709
    arrayCatchToRun(755) = 3105
    arrayCatchToRun(756) = 3195
    arrayCatchToRun(757) = 3379
    arrayCatchToRun(758) = 3254
    arrayCatchToRun(759) = 3423
    arrayCatchToRun(760) = 3276
    arrayCatchToRun(761) = 3447
    arrayCatchToRun(762) = 3298
    arrayCatchToRun(763) = 3521
    arrayCatchToRun(764) = 3623
    arrayCatchToRun(765) = 3539
    arrayCatchToRun(766) = 3830
    arrayCatchToRun(767) = 3472
    arrayCatchToRun(768) = 3648
    arrayCatchToRun(769) = 3795
    arrayCatchToRun(770) = 3744
    arrayCatchToRun(771) = 3789
    arrayCatchToRun(772) = 3848
    arrayCatchToRun(773) = 3930
    arrayCatchToRun(774) = 3607
    arrayCatchToRun(775) = 3955
    arrayCatchToRun(776) = 3814
    arrayCatchToRun(777) = 3941
    arrayCatchToRun(778) = 3976
    arrayCatchToRun(779) = 3880
    arrayCatchToRun(780) = 3986
    arrayCatchToRun(781) = 4086
    arrayCatchToRun(782) = 4055
    arrayCatchToRun(783) = 4050
    arrayCatchToRun(784) = 4628
    arrayCatchToRun(785) = 4365
    arrayCatchToRun(786) = 4729
    arrayCatchToRun(787) = 4096
    arrayCatchToRun(788) = 4142
    arrayCatchToRun(789) = 4198
    arrayCatchToRun(790) = 3933
    arrayCatchToRun(791) = 4164
    arrayCatchToRun(792) = 4587
    arrayCatchToRun(793) = 4443
    arrayCatchToRun(794) = 4461
    arrayCatchToRun(795) = 4437
    arrayCatchToRun(796) = 4156
    arrayCatchToRun(797) = 4429
    arrayCatchToRun(798) = 4076
    arrayCatchToRun(799) = 5170
    arrayCatchToRun(800) = 4520
    arrayCatchToRun(801) = 4579
    arrayCatchToRun(802) = 5075
    arrayCatchToRun(803) = 4500
    arrayCatchToRun(804) = 4668
    arrayCatchToRun(805) = 4484
    arrayCatchToRun(806) = 7330
    arrayCatchToRun(807) = 5038
    arrayCatchToRun(808) = 4749
    arrayCatchToRun(809) = 4698
    arrayCatchToRun(810) = 4710
    arrayCatchToRun(811) = 4890
    arrayCatchToRun(812) = 4494
    arrayCatchToRun(813) = 4308
    arrayCatchToRun(814) = 5350
    arrayCatchToRun(815) = 5313
    arrayCatchToRun(816) = 4880
    arrayCatchToRun(817) = 4841
    arrayCatchToRun(818) = 4444
    arrayCatchToRun(819) = 5756
    arrayCatchToRun(820) = 5435
    arrayCatchToRun(821) = 5377
    arrayCatchToRun(822) = 6892
    arrayCatchToRun(823) = 5630
    arrayCatchToRun(824) = 5714
    arrayCatchToRun(825) = 6176
    arrayCatchToRun(826) = 5885
    arrayCatchToRun(827) = 5763
    arrayCatchToRun(828) = 6609
    arrayCatchToRun(829) = 6455
    arrayCatchToRun(830) = 5891
    arrayCatchToRun(831) = 6242
    arrayCatchToRun(832) = 6432
    arrayCatchToRun(833) = 6204
    arrayCatchToRun(834) = 6222
    arrayCatchToRun(835) = 6088
    arrayCatchToRun(836) = 6077
    arrayCatchToRun(837) = 5284
    arrayCatchToRun(838) = 6119
    arrayCatchToRun(839) = 7197
    arrayCatchToRun(840) = 11837
    arrayCatchToRun(841) = 6155
    arrayCatchToRun(842) = 6476
    arrayCatchToRun(843) = 6313
    arrayCatchToRun(844) = 6393
    arrayCatchToRun(845) = 6967
    arrayCatchToRun(846) = 7308
    arrayCatchToRun(847) = 6594
    arrayCatchToRun(848) = 6674
    arrayCatchToRun(849) = 6720
    arrayCatchToRun(850) = 6924
    arrayCatchToRun(851) = 6814
    arrayCatchToRun(852) = 6804
    arrayCatchToRun(853) = 6995
    arrayCatchToRun(854) = 6633
    arrayCatchToRun(855) = 6803
    arrayCatchToRun(856) = 6685
    arrayCatchToRun(857) = 6858
    arrayCatchToRun(858) = 6976
    arrayCatchToRun(859) = 6843
    arrayCatchToRun(860) = 6921
    arrayCatchToRun(861) = 6979
    arrayCatchToRun(862) = 7191
    arrayCatchToRun(863) = 7406
    arrayCatchToRun(864) = 7233
    arrayCatchToRun(865) = 7186
    arrayCatchToRun(866) = 7353
    arrayCatchToRun(867) = 7345
    arrayCatchToRun(868) = 7302
    arrayCatchToRun(869) = 7285
    arrayCatchToRun(870) = 7257
    arrayCatchToRun(871) = 7299
    arrayCatchToRun(872) = 7331
    arrayCatchToRun(873) = 7407
    arrayCatchToRun(874) = 7961
    arrayCatchToRun(875) = 7599
    arrayCatchToRun(876) = 7554
    arrayCatchToRun(877) = 7459
    arrayCatchToRun(878) = 7322
    arrayCatchToRun(879) = 7557
    arrayCatchToRun(880) = 7469
    arrayCatchToRun(881) = 9724
    arrayCatchToRun(882) = 7802
    arrayCatchToRun(883) = 7540
    arrayCatchToRun(884) = 7792
    arrayCatchToRun(885) = 7826
    arrayCatchToRun(886) = 7740
    arrayCatchToRun(887) = 7563
    arrayCatchToRun(888) = 7622
    arrayCatchToRun(889) = 7686
    arrayCatchToRun(890) = 7831
    arrayCatchToRun(891) = 7888
    arrayCatchToRun(892) = 8168
    arrayCatchToRun(893) = 7858
    arrayCatchToRun(894) = 8704
    arrayCatchToRun(895) = 7860
    arrayCatchToRun(896) = 8189
    arrayCatchToRun(897) = 11864
    arrayCatchToRun(898) = 7991
    arrayCatchToRun(899) = 9382
    arrayCatchToRun(900) = 8560
    arrayCatchToRun(901) = 8453
    arrayCatchToRun(902) = 8298
    arrayCatchToRun(903) = 7986
    arrayCatchToRun(904) = 9151
    arrayCatchToRun(905) = 8278
    arrayCatchToRun(906) = 8355
    arrayCatchToRun(907) = 8279
    arrayCatchToRun(908) = 8778
    arrayCatchToRun(909) = 9664
    arrayCatchToRun(910) = 8570
    arrayCatchToRun(911) = 8758
    arrayCatchToRun(912) = 8447
    arrayCatchToRun(913) = 8765
    arrayCatchToRun(914) = 8739
    arrayCatchToRun(915) = 8841
    arrayCatchToRun(916) = 8899
    arrayCatchToRun(917) = 8785
    arrayCatchToRun(918) = 8830
    arrayCatchToRun(919) = 8791
    arrayCatchToRun(920) = 9011
    arrayCatchToRun(921) = 8925
    arrayCatchToRun(922) = 9567
    arrayCatchToRun(923) = 9100
    arrayCatchToRun(924) = 9029
    arrayCatchToRun(925) = 8968
    arrayCatchToRun(926) = 9251
    arrayCatchToRun(927) = 9238
    arrayCatchToRun(928) = 9569
    arrayCatchToRun(929) = 9852
    arrayCatchToRun(930) = 9429
    arrayCatchToRun(931) = 9278
    arrayCatchToRun(932) = 9425
    arrayCatchToRun(933) = 9236
    arrayCatchToRun(934) = 9593
    arrayCatchToRun(935) = 9669
    arrayCatchToRun(936) = 9485
    arrayCatchToRun(937) = 9713
    arrayCatchToRun(938) = 9738
    arrayCatchToRun(939) = 9748
    arrayCatchToRun(940) = 9922
    arrayCatchToRun(941) = 10120
    arrayCatchToRun(942) = 9827
    arrayCatchToRun(943) = 9919
    arrayCatchToRun(944) = 10156
    arrayCatchToRun(945) = 10069
    arrayCatchToRun(946) = 9995
    arrayCatchToRun(947) = 9970
    arrayCatchToRun(948) = 11338
    arrayCatchToRun(949) = 10171
    arrayCatchToRun(950) = 10159
    arrayCatchToRun(951) = 10301
    arrayCatchToRun(952) = 10158
    arrayCatchToRun(953) = 10315
    arrayCatchToRun(954) = 10283
    arrayCatchToRun(955) = 10337
    arrayCatchToRun(956) = 10389
    arrayCatchToRun(957) = 10340
    arrayCatchToRun(958) = 10387
    arrayCatchToRun(959) = 10495
    arrayCatchToRun(960) = 10417
    arrayCatchToRun(961) = 10468
    arrayCatchToRun(962) = 10535
    arrayCatchToRun(963) = 10572
    arrayCatchToRun(964) = 10633
    arrayCatchToRun(965) = 10749
    arrayCatchToRun(966) = 10743
    arrayCatchToRun(967) = 10677
    arrayCatchToRun(968) = 10709
    arrayCatchToRun(969) = 11074
    arrayCatchToRun(970) = 10724
    arrayCatchToRun(971) = 10680
    arrayCatchToRun(972) = 10606
    arrayCatchToRun(973) = 10840
    arrayCatchToRun(974) = 10856
    arrayCatchToRun(975) = 10870
    arrayCatchToRun(976) = 11100
    arrayCatchToRun(977) = 10945
    arrayCatchToRun(978) = 11232
    arrayCatchToRun(979) = 11215
    arrayCatchToRun(980) = 11070
    arrayCatchToRun(981) = 11008
    arrayCatchToRun(982) = 11211
    arrayCatchToRun(983) = 11302
    arrayCatchToRun(984) = 11157
    arrayCatchToRun(985) = 11089
    arrayCatchToRun(986) = 11095
    arrayCatchToRun(987) = 11197
    arrayCatchToRun(988) = 11228
    arrayCatchToRun(989) = 11261
    arrayCatchToRun(990) = 11073
    arrayCatchToRun(991) = 11111
    arrayCatchToRun(992) = 11175
    arrayCatchToRun(993) = 11387
    arrayCatchToRun(994) = 11643
    arrayCatchToRun(995) = 11304
    arrayCatchToRun(996) = 11461
    arrayCatchToRun(997) = 11593
    arrayCatchToRun(998) = 12368
    arrayCatchToRun(999) = 11592
    arrayCatchToRun(1000) = 11573
    arrayCatchToRun(1001) = 11626
    arrayCatchToRun(1002) = 11678
    arrayCatchToRun(1003) = 11774
    arrayCatchToRun(1004) = 11634
    arrayCatchToRun(1005) = 11916
    arrayCatchToRun(1006) = 11776
    arrayCatchToRun(1007) = 11892
    arrayCatchToRun(1008) = 11870
    arrayCatchToRun(1009) = 11862
    arrayCatchToRun(1010) = 11873
    arrayCatchToRun(1011) = 11927
    arrayCatchToRun(1012) = 11991
    arrayCatchToRun(1013) = 12065
    arrayCatchToRun(1014) = 12019
    arrayCatchToRun(1015) = 12020
    arrayCatchToRun(1016) = 12239
    arrayCatchToRun(1017) = 12040
    arrayCatchToRun(1018) = 12063
    arrayCatchToRun(1019) = 12156
    arrayCatchToRun(1020) = 12222
    arrayCatchToRun(1021) = 12073
    arrayCatchToRun(1022) = 12212
    arrayCatchToRun(1023) = 12227
    arrayCatchToRun(1024) = 12249
    arrayCatchToRun(1025) = 12289
    arrayCatchToRun(1026) = 12268
    arrayCatchToRun(1027) = 12266
    arrayCatchToRun(1028) = 13470
    arrayCatchToRun(1029) = 12310
    arrayCatchToRun(1030) = 12281
    arrayCatchToRun(1031) = 12560
    arrayCatchToRun(1032) = 12346
    arrayCatchToRun(1033) = 12414
    arrayCatchToRun(1034) = 12413
    arrayCatchToRun(1035) = 12606
    arrayCatchToRun(1036) = 12525
    arrayCatchToRun(1037) = 12495
    arrayCatchToRun(1038) = 12554
    arrayCatchToRun(1039) = 12473
    arrayCatchToRun(1040) = 12512
    arrayCatchToRun(1041) = 12539
    arrayCatchToRun(1042) = 12654
    arrayCatchToRun(1043) = 12552
    arrayCatchToRun(1044) = 12544
    arrayCatchToRun(1045) = 12611
    arrayCatchToRun(1046) = 12522
    arrayCatchToRun(1047) = 12784
    arrayCatchToRun(1048) = 12845
    arrayCatchToRun(1049) = 12637
    arrayCatchToRun(1050) = 12744
    arrayCatchToRun(1051) = 12762
    arrayCatchToRun(1052) = 12832
    arrayCatchToRun(1053) = 12791
    arrayCatchToRun(1054) = 12801
    arrayCatchToRun(1055) = 12721
    arrayCatchToRun(1056) = 12802
    arrayCatchToRun(1057) = 12695
    arrayCatchToRun(1058) = 13311
    arrayCatchToRun(1059) = 12952
    arrayCatchToRun(1060) = 12905
    arrayCatchToRun(1061) = 12823
    arrayCatchToRun(1062) = 13389
    arrayCatchToRun(1063) = 13445
    arrayCatchToRun(1064) = 12993
    arrayCatchToRun(1065) = 13037
    arrayCatchToRun(1066) = 12990
    arrayCatchToRun(1067) = 13034
    arrayCatchToRun(1068) = 13119
    arrayCatchToRun(1069) = 12956
    arrayCatchToRun(1070) = 13019
    arrayCatchToRun(1071) = 13068
    arrayCatchToRun(1072) = 13185
    arrayCatchToRun(1073) = 13013
    arrayCatchToRun(1074) = 13465
    arrayCatchToRun(1075) = 13272
    arrayCatchToRun(1076) = 13221
    arrayCatchToRun(1077) = 13237
    arrayCatchToRun(1078) = 13316
    arrayCatchToRun(1079) = 13065
    arrayCatchToRun(1080) = 13437
    arrayCatchToRun(1081) = 13359
    arrayCatchToRun(1082) = 13293
    arrayCatchToRun(1083) = 13502
    arrayCatchToRun(1084) = 13457
    arrayCatchToRun(1085) = 13524
    arrayCatchToRun(1086) = 13463
    arrayCatchToRun(1087) = 13487
    arrayCatchToRun(1088) = 13613
    arrayCatchToRun(1089) = 13419
    arrayCatchToRun(1090) = 13563
    arrayCatchToRun(1091) = 13596
    arrayCatchToRun(1092) = 13608
    arrayCatchToRun(1093) = 13705
    arrayCatchToRun(1094) = 13661
    arrayCatchToRun(1095) = 13700
    arrayCatchToRun(1096) = 13774
    arrayCatchToRun(1097) = 13746
    arrayCatchToRun(1098) = 13815
    arrayCatchToRun(1099) = 13814
    arrayCatchToRun(1100) = 13792
    arrayCatchToRun(1101) = 13899
    arrayCatchToRun(1102) = 13836
    arrayCatchToRun(1103) = 13803
    arrayCatchToRun(1104) = 16318
    arrayCatchToRun(1105) = 13696
    arrayCatchToRun(1106) = 14095
    arrayCatchToRun(1107) = 14057
    arrayCatchToRun(1108) = 13895
    arrayCatchToRun(1109) = 13979
    arrayCatchToRun(1110) = 14018
    arrayCatchToRun(1111) = 13840
    arrayCatchToRun(1112) = 13831
    arrayCatchToRun(1113) = 14049
    arrayCatchToRun(1114) = 14202
    arrayCatchToRun(1115) = 14076
    arrayCatchToRun(1116) = 14135
    arrayCatchToRun(1117) = 14362
    arrayCatchToRun(1118) = 14293
    arrayCatchToRun(1119) = 14207
    arrayCatchToRun(1120) = 14186
    arrayCatchToRun(1121) = 14229
    arrayCatchToRun(1122) = 14252
    arrayCatchToRun(1123) = 14254
    arrayCatchToRun(1124) = 14288
    arrayCatchToRun(1125) = 14318
    arrayCatchToRun(1126) = 14345
    arrayCatchToRun(1127) = 14319
    arrayCatchToRun(1128) = 14328
    arrayCatchToRun(1129) = 14385
    arrayCatchToRun(1130) = 14585
    arrayCatchToRun(1131) = 14392
    arrayCatchToRun(1132) = 14402
    arrayCatchToRun(1133) = 14491
    arrayCatchToRun(1134) = 14384
    arrayCatchToRun(1135) = 14502
    arrayCatchToRun(1136) = 14572
    arrayCatchToRun(1137) = 15007
    arrayCatchToRun(1138) = 14622
    arrayCatchToRun(1139) = 14462
    arrayCatchToRun(1140) = 14595
    arrayCatchToRun(1141) = 14735
    arrayCatchToRun(1142) = 14668
    arrayCatchToRun(1143) = 14760
    arrayCatchToRun(1144) = 15015
    arrayCatchToRun(1145) = 14811
    arrayCatchToRun(1146) = 14728
    arrayCatchToRun(1147) = 15108
    arrayCatchToRun(1148) = 15194
    arrayCatchToRun(1149) = 16661
    arrayCatchToRun(1150) = 15673
    arrayCatchToRun(1151) = 15477
    arrayCatchToRun(1152) = 15157
    arrayCatchToRun(1153) = 15332
    arrayCatchToRun(1154) = 15491
    arrayCatchToRun(1155) = 15240
    arrayCatchToRun(1156) = 15349
    arrayCatchToRun(1157) = 15657
    arrayCatchToRun(1158) = 15476
    arrayCatchToRun(1159) = 15255
    arrayCatchToRun(1160) = 15711
    arrayCatchToRun(1161) = 15634
    arrayCatchToRun(1162) = 15652
    arrayCatchToRun(1163) = 15649
    arrayCatchToRun(1164) = 15759
    arrayCatchToRun(1165) = 15720
    arrayCatchToRun(1166) = 15890
    arrayCatchToRun(1167) = 15789
    arrayCatchToRun(1168) = 15864
    arrayCatchToRun(1169) = 15862
    arrayCatchToRun(1170) = 15935
    arrayCatchToRun(1171) = 15814
    arrayCatchToRun(1172) = 15855
    arrayCatchToRun(1173) = 15962
    arrayCatchToRun(1174) = 16043
    arrayCatchToRun(1175) = 16220
    arrayCatchToRun(1176) = 16098
    arrayCatchToRun(1177) = 16239
    arrayCatchToRun(1178) = 16308
    arrayCatchToRun(1179) = 16285
    arrayCatchToRun(1180) = 16265
    arrayCatchToRun(1181) = 16859
    arrayCatchToRun(1182) = 16399
    arrayCatchToRun(1183) = 16275
    arrayCatchToRun(1184) = 16331
    arrayCatchToRun(1185) = 16258
    arrayCatchToRun(1186) = 16446
    arrayCatchToRun(1187) = 16505
    arrayCatchToRun(1188) = 16460
    arrayCatchToRun(1189) = 16469
    arrayCatchToRun(1190) = 16473
    arrayCatchToRun(1191) = 16531
    arrayCatchToRun(1192) = 16588
    arrayCatchToRun(1193) = 16587
    arrayCatchToRun(1194) = 16616
    arrayCatchToRun(1195) = 16624
    arrayCatchToRun(1196) = 16888
    arrayCatchToRun(1197) = 16794
    arrayCatchToRun(1198) = 16619
    arrayCatchToRun(1199) = 16799
    arrayCatchToRun(1200) = 18268
    arrayCatchToRun(1201) = 16766
    arrayCatchToRun(1202) = 16796
    arrayCatchToRun(1203) = 16110
    arrayCatchToRun(1204) = 16834
    arrayCatchToRun(1205) = 16966
    arrayCatchToRun(1206) = 17183
    arrayCatchToRun(1207) = 17330
    arrayCatchToRun(1208) = 17351
    arrayCatchToRun(1209) = 16831
    arrayCatchToRun(1210) = 16994
    arrayCatchToRun(1211) = 17241
    arrayCatchToRun(1212) = 17220
    arrayCatchToRun(1213) = 17150
    arrayCatchToRun(1214) = 17008
    arrayCatchToRun(1215) = 17116
    arrayCatchToRun(1216) = 17151
    arrayCatchToRun(1217) = 17596
    arrayCatchToRun(1218) = 17209
    arrayCatchToRun(1219) = 16986
    arrayCatchToRun(1220) = 17264
    arrayCatchToRun(1221) = 17355
    arrayCatchToRun(1222) = 17093
    arrayCatchToRun(1223) = 17486
    arrayCatchToRun(1224) = 16902
    arrayCatchToRun(1225) = 17447
    arrayCatchToRun(1226) = 17497
    arrayCatchToRun(1227) = 17480
    arrayCatchToRun(1228) = 17629
    arrayCatchToRun(1229) = 17652
    arrayCatchToRun(1230) = 17239
    arrayCatchToRun(1231) = 18644
    arrayCatchToRun(1232) = 17720
    arrayCatchToRun(1233) = 17700
    arrayCatchToRun(1234) = 17838
    arrayCatchToRun(1235) = 17741
    arrayCatchToRun(1236) = 17642
    arrayCatchToRun(1237) = 17861
    arrayCatchToRun(1238) = 17946
    arrayCatchToRun(1239) = 18024
    arrayCatchToRun(1240) = 18077
    arrayCatchToRun(1241) = 18062
    arrayCatchToRun(1242) = 18113
    arrayCatchToRun(1243) = 18189
    arrayCatchToRun(1244) = 18394
    arrayCatchToRun(1245) = 18360
    arrayCatchToRun(1246) = 18402
    arrayCatchToRun(1247) = 18241
    arrayCatchToRun(1248) = 18579
    arrayCatchToRun(1249) = 18589
    arrayCatchToRun(1250) = 18582
    arrayCatchToRun(1251) = 18962
    arrayCatchToRun(1252) = 18563
    arrayCatchToRun(1253) = 18616
    arrayCatchToRun(1254) = 18483
    arrayCatchToRun(1255) = 19235
    arrayCatchToRun(1256) = 18681
    arrayCatchToRun(1257) = 18430
    arrayCatchToRun(1258) = 18735
    arrayCatchToRun(1259) = 18798
    arrayCatchToRun(1260) = 18894
    arrayCatchToRun(1261) = 18931
    arrayCatchToRun(1262) = 18882
    arrayCatchToRun(1263) = 19148
    arrayCatchToRun(1264) = 19472
    arrayCatchToRun(1265) = 19376
    arrayCatchToRun(1266) = 19222
    arrayCatchToRun(1267) = 18825
    arrayCatchToRun(1268) = 19225
    arrayCatchToRun(1269) = 19251
    arrayCatchToRun(1270) = 19555
    arrayCatchToRun(1271) = 19335
    arrayCatchToRun(1272) = 19516
    arrayCatchToRun(1273) = 19409
    arrayCatchToRun(1274) = 19330
    arrayCatchToRun(1275) = 19574
    arrayCatchToRun(1276) = 19371
    arrayCatchToRun(1277) = 19467
    arrayCatchToRun(1278) = 19485
    arrayCatchToRun(1279) = 19584
    arrayCatchToRun(1280) = 19691
    arrayCatchToRun(1281) = 19593
    arrayCatchToRun(1282) = 19736
    arrayCatchToRun(1283) = 19572
    arrayCatchToRun(1284) = 19842
    arrayCatchToRun(1285) = 20068
    arrayCatchToRun(1286) = 20070
    arrayCatchToRun(1287) = 19886
    arrayCatchToRun(1288) = 19863
    arrayCatchToRun(1289) = 19947
    arrayCatchToRun(1290) = 20163
    arrayCatchToRun(1291) = 20043
    arrayCatchToRun(1292) = 20156
    arrayCatchToRun(1293) = 20282
    arrayCatchToRun(1294) = 20185
    arrayCatchToRun(1295) = 20304
    arrayCatchToRun(1296) = 20247
    arrayCatchToRun(1297) = 20344
    arrayCatchToRun(1298) = 20321
    arrayCatchToRun(1299) = 20342
    arrayCatchToRun(1300) = 20408
    arrayCatchToRun(1301) = 20468
    arrayCatchToRun(1302) = 20454
    arrayCatchToRun(1303) = 20476
    arrayCatchToRun(1304) = 20445
    arrayCatchToRun(1305) = 20613
    arrayCatchToRun(1306) = 20610
    arrayCatchToRun(1307) = 20580
    arrayCatchToRun(1308) = 20481
    arrayCatchToRun(1309) = 20675
    arrayCatchToRun(1310) = 20257
    arrayCatchToRun(1311) = 20656
    arrayCatchToRun(1312) = 20829
    arrayCatchToRun(1313) = 20736
    arrayCatchToRun(1314) = 20862
    arrayCatchToRun(1315) = 20916
    arrayCatchToRun(1316) = 20861
    arrayCatchToRun(1317) = 21123
    arrayCatchToRun(1318) = 21026
    arrayCatchToRun(1319) = 21130
    arrayCatchToRun(1320) = 21099
    arrayCatchToRun(1321) = 21364
    arrayCatchToRun(1322) = 21238
    arrayCatchToRun(1323) = 21152
    arrayCatchToRun(1324) = 21208
    arrayCatchToRun(1325) = 21224
    arrayCatchToRun(1326) = 21234
    arrayCatchToRun(1327) = 21278
    arrayCatchToRun(1328) = 22038
    arrayCatchToRun(1329) = 21466
    arrayCatchToRun(1330) = 21411
    arrayCatchToRun(1331) = 21481
    arrayCatchToRun(1332) = 21406
    arrayCatchToRun(1333) = 21737
    arrayCatchToRun(1334) = 21638
    arrayCatchToRun(1335) = 21701
    arrayCatchToRun(1336) = 21763
    arrayCatchToRun(1337) = 22725
    arrayCatchToRun(1338) = 21774
    arrayCatchToRun(1339) = 21839
    arrayCatchToRun(1340) = 21790
    arrayCatchToRun(1341) = 21925
    arrayCatchToRun(1342) = 21877
    arrayCatchToRun(1343) = 21793
    arrayCatchToRun(1344) = 22011
    arrayCatchToRun(1345) = 21945
    arrayCatchToRun(1346) = 23626
    arrayCatchToRun(1347) = 22006
    arrayCatchToRun(1348) = 22033
    arrayCatchToRun(1349) = 22066
    arrayCatchToRun(1350) = 22031
    arrayCatchToRun(1351) = 22179
    arrayCatchToRun(1352) = 22074
    arrayCatchToRun(1353) = 22108
    arrayCatchToRun(1354) = 22123
    arrayCatchToRun(1355) = 22103
    arrayCatchToRun(1356) = 22230
    arrayCatchToRun(1357) = 22241
    arrayCatchToRun(1358) = 22419
    arrayCatchToRun(1359) = 22191
    arrayCatchToRun(1360) = 22287
    arrayCatchToRun(1361) = 22369
    arrayCatchToRun(1362) = 22577
    arrayCatchToRun(1363) = 22449
    arrayCatchToRun(1364) = 22700
    arrayCatchToRun(1365) = 22368
    arrayCatchToRun(1366) = 22496
    arrayCatchToRun(1367) = 22468
    arrayCatchToRun(1368) = 22513
    arrayCatchToRun(1369) = 22363
    arrayCatchToRun(1370) = 22599
    arrayCatchToRun(1371) = 22647
    arrayCatchToRun(1372) = 22616
    arrayCatchToRun(1373) = 22673
    arrayCatchToRun(1374) = 22606
    arrayCatchToRun(1375) = 22680
    arrayCatchToRun(1376) = 22762
    arrayCatchToRun(1377) = 22774
    arrayCatchToRun(1378) = 22847
    arrayCatchToRun(1379) = 22790
    arrayCatchToRun(1380) = 22920
    arrayCatchToRun(1381) = 22924
    arrayCatchToRun(1382) = 22943
    arrayCatchToRun(1383) = 23274
    arrayCatchToRun(1384) = 23502
    arrayCatchToRun(1385) = 23133
    arrayCatchToRun(1386) = 23194
    arrayCatchToRun(1387) = 23108
    arrayCatchToRun(1388) = 23418
    arrayCatchToRun(1389) = 23236
    arrayCatchToRun(1390) = 23173
    arrayCatchToRun(1391) = 23216
    arrayCatchToRun(1392) = 23459
    arrayCatchToRun(1393) = 24025
    arrayCatchToRun(1394) = 23553
    arrayCatchToRun(1395) = 23472
    arrayCatchToRun(1396) = 23525
    arrayCatchToRun(1397) = 23445
    arrayCatchToRun(1398) = 23465
    arrayCatchToRun(1399) = 23555
    arrayCatchToRun(1400) = 23586
    arrayCatchToRun(1401) = 23618
    arrayCatchToRun(1402) = 23601
    arrayCatchToRun(1403) = 23515
    arrayCatchToRun(1404) = 23634
    arrayCatchToRun(1405) = 23697
    arrayCatchToRun(1406) = 23684
    arrayCatchToRun(1407) = 23770
    arrayCatchToRun(1408) = 23703
    arrayCatchToRun(1409) = 23711
    arrayCatchToRun(1410) = 23808
    arrayCatchToRun(1411) = 23921
    arrayCatchToRun(1412) = 23968
    arrayCatchToRun(1413) = 23973
    arrayCatchToRun(1414) = 24249
    arrayCatchToRun(1415) = 24016
    arrayCatchToRun(1416) = 23931
    arrayCatchToRun(1417) = 24081
    arrayCatchToRun(1418) = 24057
    arrayCatchToRun(1419) = 24124
    arrayCatchToRun(1420) = 24447
    arrayCatchToRun(1421) = 24729
    arrayCatchToRun(1422) = 24171
    arrayCatchToRun(1423) = 24308
    arrayCatchToRun(1424) = 24230
    arrayCatchToRun(1425) = 24292
    arrayCatchToRun(1426) = 24297
    arrayCatchToRun(1427) = 24358
    arrayCatchToRun(1428) = 24438
    arrayCatchToRun(1429) = 24519
    arrayCatchToRun(1430) = 24450
    arrayCatchToRun(1431) = 24758
    arrayCatchToRun(1432) = 24509
    arrayCatchToRun(1433) = 24502
    arrayCatchToRun(1434) = 24654
    arrayCatchToRun(1435) = 24699
    arrayCatchToRun(1436) = 24910
    arrayCatchToRun(1437) = 24775
    arrayCatchToRun(1438) = 24843
    arrayCatchToRun(1439) = 24844
    arrayCatchToRun(1440) = 24774
    arrayCatchToRun(1441) = 24829
    arrayCatchToRun(1442) = 24815
    arrayCatchToRun(1443) = 24840
    arrayCatchToRun(1444) = 24859
    arrayCatchToRun(1445) = 24888
    arrayCatchToRun(1446) = 24899
    arrayCatchToRun(1447) = 24985
    arrayCatchToRun(1448) = 24999
    arrayCatchToRun(1449) = 25369
    arrayCatchToRun(1450) = 25001
    arrayCatchToRun(1451) = 25005
    arrayCatchToRun(1452) = 25006
    arrayCatchToRun(1453) = 25009
    arrayCatchToRun(1454) = 25048
    arrayCatchToRun(1455) = 25046
    arrayCatchToRun(1456) = 25226
    arrayCatchToRun(1457) = 25104
    arrayCatchToRun(1458) = 25166
    arrayCatchToRun(1459) = 25179
    arrayCatchToRun(1460) = 25145
    arrayCatchToRun(1461) = 25188
    arrayCatchToRun(1462) = 25184
    arrayCatchToRun(1463) = 25165
    arrayCatchToRun(1464) = 25247
    arrayCatchToRun(1465) = 25228
    arrayCatchToRun(1466) = 25277
    arrayCatchToRun(1467) = 25280
    arrayCatchToRun(1468) = 25313
    arrayCatchToRun(1469) = 25266
    arrayCatchToRun(1470) = 25352
    arrayCatchToRun(1471) = 25365
    arrayCatchToRun(1472) = 25396
    arrayCatchToRun(1473) = 25394
    arrayCatchToRun(1474) = 25370
    arrayCatchToRun(1475) = 25368
    arrayCatchToRun(1476) = 25497
    arrayCatchToRun(1477) = 25464
    arrayCatchToRun(1478) = 25476
    arrayCatchToRun(1479) = 25499
    arrayCatchToRun(1480) = 25487
    arrayCatchToRun(1481) = 25502
    arrayCatchToRun(1482) = 25532
    arrayCatchToRun(1483) = 25548
    arrayCatchToRun(1484) = 25553
    arrayCatchToRun(1485) = 25595
    arrayCatchToRun(1486) = 25589
    arrayCatchToRun(1487) = 25710
    arrayCatchToRun(1488) = 25652
    arrayCatchToRun(1489) = 25616
    arrayCatchToRun(1490) = 25650
    arrayCatchToRun(1491) = 25786
    arrayCatchToRun(1492) = 25597
    arrayCatchToRun(1493) = 25855
    arrayCatchToRun(1494) = 25735
    arrayCatchToRun(1495) = 25705
    arrayCatchToRun(1496) = 25793
    arrayCatchToRun(1497) = 26139
    arrayCatchToRun(1498) = 25834
    arrayCatchToRun(1499) = 25870
    arrayCatchToRun(1500) = 25873
    arrayCatchToRun(1501) = 25908
    arrayCatchToRun(1502) = 25872
    arrayCatchToRun(1503) = 25864
    arrayCatchToRun(1504) = 25950
    arrayCatchToRun(1505) = 25930
    arrayCatchToRun(1506) = 25862
    arrayCatchToRun(1507) = 25899
    arrayCatchToRun(1508) = 25957
    arrayCatchToRun(1509) = 26114
    arrayCatchToRun(1510) = 26003
    arrayCatchToRun(1511) = 26168
    arrayCatchToRun(1512) = 26016
    arrayCatchToRun(1513) = 26163
    arrayCatchToRun(1514) = 26257
    arrayCatchToRun(1515) = 26176
    arrayCatchToRun(1516) = 26178
    arrayCatchToRun(1517) = 25737
    arrayCatchToRun(1518) = 26247
    arrayCatchToRun(1519) = 26392
    arrayCatchToRun(1520) = 26244
    arrayCatchToRun(1521) = 26274
    arrayCatchToRun(1522) = 26260
    arrayCatchToRun(1523) = 26275
    arrayCatchToRun(1524) = 26459
    arrayCatchToRun(1525) = 26456
    arrayCatchToRun(1526) = 26455
    arrayCatchToRun(1527) = 26480
    arrayCatchToRun(1528) = 26520
    arrayCatchToRun(1529) = 26527
    arrayCatchToRun(1530) = 26550
    arrayCatchToRun(1531) = 26518
    arrayCatchToRun(1532) = 26529
    arrayCatchToRun(1533) = 26493
    arrayCatchToRun(1534) = 26578
    arrayCatchToRun(1535) = 26671
    arrayCatchToRun(1536) = 26617
    arrayCatchToRun(1537) = 26596
    arrayCatchToRun(1538) = 26637
    arrayCatchToRun(1539) = 26732
    arrayCatchToRun(1540) = 26692
    arrayCatchToRun(1541) = 26700
    arrayCatchToRun(1542) = 26440
    arrayCatchToRun(1543) = 26729
    arrayCatchToRun(1544) = 26735
    arrayCatchToRun(1545) = 26773
    arrayCatchToRun(1546) = 26787
    arrayCatchToRun(1547) = 26795
    arrayCatchToRun(1548) = 26872
    arrayCatchToRun(1549) = 26845
    arrayCatchToRun(1550) = 26895
    arrayCatchToRun(1551) = 26883
    arrayCatchToRun(1552) = 27043
    arrayCatchToRun(1553) = 26944
    arrayCatchToRun(1554) = 27067
    arrayCatchToRun(1555) = 27080
    arrayCatchToRun(1556) = 27220
    arrayCatchToRun(1557) = 27247
    arrayCatchToRun(1558) = 27249
    arrayCatchToRun(1559) = 27331
    arrayCatchToRun(1560) = 27340
    arrayCatchToRun(1561) = 27373
    arrayCatchToRun(1562) = 27398
    arrayCatchToRun(1563) = 27435
    arrayCatchToRun(1564) = 27322
    arrayCatchToRun(1565) = 27454
    arrayCatchToRun(1566) = 27553
    arrayCatchToRun(1567) = 27586
    arrayCatchToRun(1568) = 27638
    arrayCatchToRun(1569) = 28003
    arrayCatchToRun(1570) = 27936
    arrayCatchToRun(1571) = 27899
    arrayCatchToRun(1572) = 28014
    arrayCatchToRun(1573) = 27996
    arrayCatchToRun(1574) = 28043
    arrayCatchToRun(1575) = 28056
    arrayCatchToRun(1576) = 28071
    arrayCatchToRun(1577) = 28115
    arrayCatchToRun(1578) = 28178
    arrayCatchToRun(1579) = 28158
    arrayCatchToRun(1580) = 28199
    arrayCatchToRun(1581) = 28188
    arrayCatchToRun(1582) = 28268
    arrayCatchToRun(1583) = 28289
    arrayCatchToRun(1584) = 28303
    arrayCatchToRun(1585) = 28323
    arrayCatchToRun(1586) = 28342
    arrayCatchToRun(1587) = 28384
    arrayCatchToRun(1588) = 28461
    arrayCatchToRun(1589) = 28416
    arrayCatchToRun(1590) = 28426
    arrayCatchToRun(1591) = 28487
    arrayCatchToRun(1592) = 28478
    arrayCatchToRun(1593) = 28506
End If


If intArrayToUse = 4 Then
    ReDim arrayCatchToRun(332) As Integer
    arrayCatchToRun(0) = 7
    arrayCatchToRun(1) = 704
    arrayCatchToRun(2) = 1135
    arrayCatchToRun(3) = 1271
    arrayCatchToRun(4) = 1570
    arrayCatchToRun(5) = 1678
    arrayCatchToRun(6) = 1692
    arrayCatchToRun(7) = 1694
    arrayCatchToRun(8) = 1753
    arrayCatchToRun(9) = 1764
    arrayCatchToRun(10) = 1847
    arrayCatchToRun(11) = 1853
    arrayCatchToRun(12) = 2035
    arrayCatchToRun(13) = 2098
    arrayCatchToRun(14) = 2139
    arrayCatchToRun(15) = 2237
    arrayCatchToRun(16) = 2358
    arrayCatchToRun(17) = 2402
    arrayCatchToRun(18) = 2490
    arrayCatchToRun(19) = 2499
    arrayCatchToRun(20) = 2580
    arrayCatchToRun(21) = 2712
    arrayCatchToRun(22) = 2831
    arrayCatchToRun(23) = 3105
    arrayCatchToRun(24) = 3458
    arrayCatchToRun(25) = 3527
    arrayCatchToRun(26) = 3532
    arrayCatchToRun(27) = 3636
    arrayCatchToRun(28) = 3904
    arrayCatchToRun(29) = 3930
    arrayCatchToRun(30) = 4164
    arrayCatchToRun(31) = 4229
    arrayCatchToRun(32) = 4284
    arrayCatchToRun(33) = 4444
    arrayCatchToRun(34) = 4672
    arrayCatchToRun(35) = 4796
    arrayCatchToRun(36) = 5222
    arrayCatchToRun(37) = 5284
    arrayCatchToRun(38) = 5350
    arrayCatchToRun(39) = 5605
    arrayCatchToRun(40) = 6140
    arrayCatchToRun(41) = 6297
    arrayCatchToRun(42) = 6405
    arrayCatchToRun(43) = 6455
    arrayCatchToRun(44) = 6687
    arrayCatchToRun(45) = 7092
    arrayCatchToRun(46) = 7183
    arrayCatchToRun(47) = 7222
    arrayCatchToRun(48) = 7469
    arrayCatchToRun(49) = 7588
    arrayCatchToRun(50) = 7691
    arrayCatchToRun(51) = 7730
    arrayCatchToRun(52) = 8090
    arrayCatchToRun(53) = 8097
    arrayCatchToRun(54) = 8168
    arrayCatchToRun(55) = 8307
    arrayCatchToRun(56) = 8361
    arrayCatchToRun(57) = 8455
    arrayCatchToRun(58) = 8564
    arrayCatchToRun(59) = 8738
    arrayCatchToRun(60) = 8751
    arrayCatchToRun(61) = 8937
    arrayCatchToRun(62) = 8959
    arrayCatchToRun(63) = 8971
    arrayCatchToRun(64) = 9048
    arrayCatchToRun(65) = 9098
    arrayCatchToRun(66) = 9176
    arrayCatchToRun(67) = 9281
    arrayCatchToRun(68) = 9351
    arrayCatchToRun(69) = 9401
    arrayCatchToRun(70) = 9617
    arrayCatchToRun(71) = 9658
    arrayCatchToRun(72) = 9852
    arrayCatchToRun(73) = 9922
    arrayCatchToRun(74) = 10060
    arrayCatchToRun(75) = 10132
    arrayCatchToRun(76) = 10301
    arrayCatchToRun(77) = 10389
    arrayCatchToRun(78) = 10719
    arrayCatchToRun(79) = 10786
    arrayCatchToRun(80) = 10934
    arrayCatchToRun(81) = 11101
    arrayCatchToRun(82) = 11187
    arrayCatchToRun(83) = 11189
    arrayCatchToRun(84) = 11337
    arrayCatchToRun(85) = 11338
    arrayCatchToRun(86) = 11339
    arrayCatchToRun(87) = 11384
    arrayCatchToRun(88) = 11385
    arrayCatchToRun(89) = 11447
    arrayCatchToRun(90) = 11504
    arrayCatchToRun(91) = 11611
    arrayCatchToRun(92) = 11642
    arrayCatchToRun(93) = 12055
    arrayCatchToRun(94) = 12313
    arrayCatchToRun(95) = 12382
    arrayCatchToRun(96) = 12606
    arrayCatchToRun(97) = 12659
    arrayCatchToRun(98) = 12978
    arrayCatchToRun(99) = 12995
    arrayCatchToRun(100) = 13445
    arrayCatchToRun(101) = 13463
    arrayCatchToRun(102) = 13544
    arrayCatchToRun(103) = 13696
    arrayCatchToRun(104) = 13780
    arrayCatchToRun(105) = 13972
    arrayCatchToRun(106) = 14019
    arrayCatchToRun(107) = 14032
    arrayCatchToRun(108) = 14057
    arrayCatchToRun(109) = 14098
    arrayCatchToRun(110) = 14202
    arrayCatchToRun(111) = 14293
    arrayCatchToRun(112) = 14315
    arrayCatchToRun(113) = 14362
    arrayCatchToRun(114) = 14384
    arrayCatchToRun(115) = 14443
    arrayCatchToRun(116) = 14585
    arrayCatchToRun(117) = 14627
    arrayCatchToRun(118) = 14740
    arrayCatchToRun(119) = 14749
    arrayCatchToRun(120) = 15026
    arrayCatchToRun(121) = 15139
    arrayCatchToRun(122) = 15222
    arrayCatchToRun(123) = 15265
    arrayCatchToRun(124) = 15390
    arrayCatchToRun(125) = 15477
    arrayCatchToRun(126) = 15651
    arrayCatchToRun(127) = 15831
    arrayCatchToRun(128) = 15935
    arrayCatchToRun(129) = 16206
    arrayCatchToRun(130) = 16220
    arrayCatchToRun(131) = 16275
    arrayCatchToRun(132) = 16369
    arrayCatchToRun(133) = 16443
    arrayCatchToRun(134) = 16456
    arrayCatchToRun(135) = 16505
    arrayCatchToRun(136) = 16624
    arrayCatchToRun(137) = 16661
    arrayCatchToRun(138) = 16902
    arrayCatchToRun(139) = 16906
    arrayCatchToRun(140) = 16986
    arrayCatchToRun(141) = 17086
    arrayCatchToRun(142) = 17239
    arrayCatchToRun(143) = 17241
    arrayCatchToRun(144) = 17619
    arrayCatchToRun(145) = 18216
    arrayCatchToRun(146) = 18607
    arrayCatchToRun(147) = 18644
    arrayCatchToRun(148) = 18645
    arrayCatchToRun(149) = 18682
    arrayCatchToRun(150) = 18767
    arrayCatchToRun(151) = 18825
    arrayCatchToRun(152) = 18876
    arrayCatchToRun(153) = 18982
    arrayCatchToRun(154) = 19079
    arrayCatchToRun(155) = 19214
    arrayCatchToRun(156) = 19261
    arrayCatchToRun(157) = 19283
    arrayCatchToRun(158) = 19381
    arrayCatchToRun(159) = 19445
    arrayCatchToRun(160) = 19516
    arrayCatchToRun(161) = 19540
    arrayCatchToRun(162) = 19572
    arrayCatchToRun(163) = 19864
    arrayCatchToRun(164) = 19896
    arrayCatchToRun(165) = 19935
    arrayCatchToRun(166) = 19952
    arrayCatchToRun(167) = 20002
    arrayCatchToRun(168) = 20043
    arrayCatchToRun(169) = 20185
    arrayCatchToRun(170) = 20196
    arrayCatchToRun(171) = 20465
    arrayCatchToRun(172) = 20573
    arrayCatchToRun(173) = 20601
    arrayCatchToRun(174) = 20633
    arrayCatchToRun(175) = 20647
    arrayCatchToRun(176) = 20657
    arrayCatchToRun(177) = 20742
    arrayCatchToRun(178) = 20754
    arrayCatchToRun(179) = 20757
    arrayCatchToRun(180) = 20828
    arrayCatchToRun(181) = 20860
    arrayCatchToRun(182) = 20965
    arrayCatchToRun(183) = 20986
    arrayCatchToRun(184) = 21023
    arrayCatchToRun(185) = 21152
    arrayCatchToRun(186) = 21189
    arrayCatchToRun(187) = 21191
    arrayCatchToRun(188) = 21328
    arrayCatchToRun(189) = 21437
    arrayCatchToRun(190) = 21466
    arrayCatchToRun(191) = 21490
    arrayCatchToRun(192) = 21576
    arrayCatchToRun(193) = 21649
    arrayCatchToRun(194) = 21751
    arrayCatchToRun(195) = 21754
    arrayCatchToRun(196) = 21790
    arrayCatchToRun(197) = 21795
    arrayCatchToRun(198) = 21823
    arrayCatchToRun(199) = 21847
    arrayCatchToRun(200) = 21848
    arrayCatchToRun(201) = 21877
    arrayCatchToRun(202) = 21925
    arrayCatchToRun(203) = 21945
    arrayCatchToRun(204) = 21956
    arrayCatchToRun(205) = 22010
    arrayCatchToRun(206) = 22086
    arrayCatchToRun(207) = 22191
    arrayCatchToRun(208) = 22259
    arrayCatchToRun(209) = 22308
    arrayCatchToRun(210) = 22419
    arrayCatchToRun(211) = 22496
    arrayCatchToRun(212) = 22610
    arrayCatchToRun(213) = 22666
    arrayCatchToRun(214) = 22725
    arrayCatchToRun(215) = 22782
    arrayCatchToRun(216) = 22787
    arrayCatchToRun(217) = 22839
    arrayCatchToRun(218) = 22840
    arrayCatchToRun(219) = 22942
    arrayCatchToRun(220) = 23192
    arrayCatchToRun(221) = 23194
    arrayCatchToRun(222) = 23206
    arrayCatchToRun(223) = 23216
    arrayCatchToRun(224) = 23245
    arrayCatchToRun(225) = 23361
    arrayCatchToRun(226) = 23465
    arrayCatchToRun(227) = 23515
    arrayCatchToRun(228) = 23553
    arrayCatchToRun(229) = 23559
    arrayCatchToRun(230) = 23561
    arrayCatchToRun(231) = 23578
    arrayCatchToRun(232) = 23618
    arrayCatchToRun(233) = 23624
    arrayCatchToRun(234) = 23654
    arrayCatchToRun(235) = 23684
    arrayCatchToRun(236) = 23711
    arrayCatchToRun(237) = 23887
    arrayCatchToRun(238) = 23938
    arrayCatchToRun(239) = 23973
    arrayCatchToRun(240) = 24016
    arrayCatchToRun(241) = 24056
    arrayCatchToRun(242) = 24103
    arrayCatchToRun(243) = 24124
    arrayCatchToRun(244) = 24132
    arrayCatchToRun(245) = 24276
    arrayCatchToRun(246) = 24280
    arrayCatchToRun(247) = 24295
    arrayCatchToRun(248) = 24297
    arrayCatchToRun(249) = 24417
    arrayCatchToRun(250) = 24459
    arrayCatchToRun(251) = 24464
    arrayCatchToRun(252) = 24522
    arrayCatchToRun(253) = 24531
    arrayCatchToRun(254) = 24623
    arrayCatchToRun(255) = 24668
    arrayCatchToRun(256) = 24744
    arrayCatchToRun(257) = 24754
    arrayCatchToRun(258) = 24758
    arrayCatchToRun(259) = 24785
    arrayCatchToRun(260) = 24798
    arrayCatchToRun(261) = 24843
    arrayCatchToRun(262) = 24892
    arrayCatchToRun(263) = 24919
    arrayCatchToRun(264) = 24996
    arrayCatchToRun(265) = 25000
    arrayCatchToRun(266) = 25006
    arrayCatchToRun(267) = 25035
    arrayCatchToRun(268) = 25038
    arrayCatchToRun(269) = 25077
    arrayCatchToRun(270) = 25128
    arrayCatchToRun(271) = 25366
    arrayCatchToRun(272) = 25378
    arrayCatchToRun(273) = 25391
    arrayCatchToRun(274) = 25400
    arrayCatchToRun(275) = 25889
    arrayCatchToRun(276) = 25925
    arrayCatchToRun(277) = 26101
    arrayCatchToRun(278) = 26114
    arrayCatchToRun(279) = 26168
    arrayCatchToRun(280) = 26237
    arrayCatchToRun(281) = 26240
    arrayCatchToRun(282) = 26243
    arrayCatchToRun(283) = 26275
    arrayCatchToRun(284) = 26356
    arrayCatchToRun(285) = 26392
    arrayCatchToRun(286) = 26416
    arrayCatchToRun(287) = 26447
    arrayCatchToRun(288) = 26450
    arrayCatchToRun(289) = 26467
    arrayCatchToRun(290) = 26472
    arrayCatchToRun(291) = 26566
    arrayCatchToRun(292) = 26578
    arrayCatchToRun(293) = 26581
    arrayCatchToRun(294) = 26692
    arrayCatchToRun(295) = 26752
    arrayCatchToRun(296) = 26804
    arrayCatchToRun(297) = 27186
    arrayCatchToRun(298) = 27309
    arrayCatchToRun(299) = 27310
    arrayCatchToRun(300) = 27315
    arrayCatchToRun(301) = 27322
    arrayCatchToRun(302) = 27361
    arrayCatchToRun(303) = 27435
    arrayCatchToRun(304) = 27523
    arrayCatchToRun(305) = 27604
    arrayCatchToRun(306) = 27627
    arrayCatchToRun(307) = 27638
    arrayCatchToRun(308) = 27675
    arrayCatchToRun(309) = 27699
    arrayCatchToRun(310) = 27785
    arrayCatchToRun(311) = 27808
    arrayCatchToRun(312) = 27848
    arrayCatchToRun(313) = 27899
    arrayCatchToRun(314) = 27936
    arrayCatchToRun(315) = 27948
    arrayCatchToRun(316) = 27967
    arrayCatchToRun(317) = 28003
    arrayCatchToRun(318) = 28014
    arrayCatchToRun(319) = 28043
    arrayCatchToRun(320) = 28071
    arrayCatchToRun(321) = 28111
    arrayCatchToRun(322) = 28130
    arrayCatchToRun(323) = 28158
    arrayCatchToRun(324) = 28200
    arrayCatchToRun(325) = 28288
    arrayCatchToRun(326) = 28330
    arrayCatchToRun(327) = 28344
    arrayCatchToRun(328) = 28493
    arrayCatchToRun(329) = 28506
End If

'read in the list of lochs and process them
cboCatchment.Text = "LocalCatchment_and_Network"

FindFLayer pFLayerCatchment, cboCatchment.Text, blnFound
Set pFClassCatchment = pFLayerCatchment.FeatureClass
Set pTableCatchment = pFClassCatchment
'Debug.Print "pTableCatchment is " & cboCatchment.Text

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter

Dim lonMin As Long
Dim lonMax As Long

lonMin = 0
lonMax = lonMin + 8030

pQueryFilt.WhereClause = "OBJECTID >= " & lonMin & " AND OBJECTID <= " & lonMax

Dim pCursor As ICursor
Set pCursor = pTableCatchment.Search(pQueryFilt, False)
Dim pRow As IRow
Set pRow = pCursor.NextRow

Dim lonRowScroller As Long
lonRowScroller = 0
'null nothing
cmdLoad_Data_Click

cmdReduce_Click

'Dim blnStartProc As Boolean
'blnStartProc = False
Dim filterMin As Long
Dim filterMax As Long

filterMin = 6800
filterMax = filterMin + 1300

Dim blnIsInArray As Boolean

Dim blnContinue As Boolean
Dim strCheckName As String

While Not pRow Is Nothing
    'If pRow.Value(pTableCatchment.FindField("SITECODE")) = 804 Then
    '    blnStartProc = True
    'End If
    'check if in array
    
    blnIsInArray = False
    Dim i As Long
    For i = 0 To UBound(arrayCatchToRun, 1)
        If (arrayCatchToRun(i) = pRow.Value(pTableCatchment.FindField("SITECODE"))) Then
            blnIsInArray = True
        End If
    Next
    
   If (blnIsInArray = True) Then
    blnContinue = True
    'If (lonRowScroller >= filterMin) And (lonRowScroller <= filterMax) Then
    'If (lonRowScroller >= lonMin) And (lonRowScroller <= (lonMax - lonMin)) And (blnStartProc = True) Then
        If (fileCheckExists = "process if missing") Then
        'test names
            strCheckName = strOutputFolder & pRow.Value(pTableCatchment.FindField("SITECODE")) & "_Capacity.txt"
            If fso.FileExists(strCheckName) Then
                'Debug.Print strCheckName & " exists"
                blnContinue = False
            End If
         End If
         If (blnContinue = True) Then
            cboGBLAKES_IDs.Text = pRow.Value(pTableCatchment.FindField("SITECODE")) & " - " & pRow.Value(pTableCatchment.FindField("SiteName"))
            'Debug.Print lonRowScroller & ", " & pRow.Value(pTableCatchment.FindField("SITECODE")) & " - " & pRow.Value(pTableCatchment.FindField("SiteName"))
            DoEvents
            cmdGetCatchmentInfo_Click
            If Not runStandardLoads Then
                txtPerCapitaTPLoadUrban.Text = 0.9125
                txtPerCapitaTPLoadRural.Text = 0.25
                chkPerCapitaTPLoadUrban.Value = True
                chkPerCapitaTPLoadUrbanAll.Value = True
                chkPerCapitaTPLoadRural.Value = True
                chkPerCapitaTPLoadRuralAll.Value = True
            End If
            cmdCalcTP_Click
            txtOutputReport.Text = strOutputFolder & pRow.Value(pTableCatchment.FindField("SITECODE"))
            txtOutputFile.Text = strOutputFolder & pRow.Value(pTableCatchment.FindField("SITECODE"))
            chkProduceResultsTable.Value = False
            DoEvents
            cmdCreateReportBatch    'this has the image stuff disabled and will run for the whole database
         End If
     End If
    'clear up memory usage
    Set SupplementColumnHeaders = Nothing
    Set List_ItemSupplement = Nothing
    Set CatchmentInfoColumnHeaders1 = Nothing
    Set List_Item2 = Nothing
    Set List_Item3 = Nothing
    Set CatchmentRelateColumnHeaders2 = Nothing
    Set pGraphicsContainer = Nothing
    'go to next row
    Set pRow = pCursor.NextRow
    lonRowScroller = lonRowScroller + 1
Wend
cmdEnlarge_Click
MsgBox "Done!!!"

End Sub

Private Sub cboResolveAreaDifference_Change()
'make the strSelectedLandCoverType equal to that selected
strSelectedLandCoverType = cboResolveAreaDifference.Text
If cboResolveAreaDifference <> "" Then
    blnUseLCoverComboSelection = True
    strNewLCoverToUse = cboResolveAreaDifference.Value
    cmdResolveAreaDifference.Caption = "Resolve area difference by adjusting " & strNewLCoverToUse
End If
txtEnterNewArea.Text = Format(Abs(dblUserModifiedLCoverArea_difference), "#,#")
chkChangeP.Value = False
chkChangeArea.Caption = "Change area of " & cboResolveAreaDifference.Value
chkChangeP.Caption = "Change P export coeff. of " & cboResolveAreaDifference.Value

'only update the new P value if there is a change in the check sum area
If Abs(dblUserModifiedLCoverArea_difference) <> 0 Then
Dim i As Long
For i = 0 To UBound(varArrayExportsTable, 1)
    If cboResolveAreaDifference.Value = varArrayExportsTable(i, 3) Then
        txtEnterNewP.Value = varArrayExportsTable(i, 6)
        Exit For
    End If
Next
End If

End Sub
Private Sub cboGBLAKES_IDs_Click()
'this is designed to clear much of the settings on the interface to the tool when the user clicks to select a new catchment
cmdCalcTP.Enabled = False
cmdModCatchmentInputs.Enabled = False
cmdResetModifiedValues.Enabled = False
chkChangeP.Enabled = False
cmdInfoChangeP.Enabled = False
chkChangeArea.Enabled = False
chkChangePforNetwork.Enabled = False
lblSelectedLandCover = ""
chkChangeP.Caption = "Change P export coeff."
chkChangeArea.Caption = "Change Area"
chkChangePforNetwork.Caption = "Change P export coeff. for whole network"
txtEnterNewP.Visible = False
txtEnterNewArea.Visible = False
lvwCatchmentInfo.Visible = False
optReadIn.Visible = False
optUserModified.Visible = False
lvwCatchmentRelationships1.Height = 0
lvwCatchmentRelationships2.Height = 0
lvwCatchmentRelationships1.Visible = False
lvwCatchmentRelationships2.Visible = False
lstViewSupplement.Visible = False
optCatchment.Value = True
optNetwork.Enabled = False
lblEnterNewP.Caption = ""
lblEnterNewArea.Caption = ""
Frame_Modify_Load.Visible = False
frameModifyInputs.Visible = False
lblScenarioSaveWarning.Visible = True
lblScenarioSaveWarning2.Visible = True
lblReportSaveWarning.Visible = True
Frame1Scenario.Visible = False
Frame2UserLandCover.Visible = False
chkChangeP.Value = False
chkChangePforNetwork.Value = False
chkChangeArea.Value = False
txtEnterNewArea.Text = ""
txtEnterNewP.Text = ""
optReadIn.Value = True
optUserModified.Value = False
chkProduceResultsTable.Enabled = False
chkProduceResultsCSV.Enabled = False
chkProduceResultsTable.Value = False
chkProduceResultsCSV.Value = False
cmdZoomToSelected.Enabled = False
cmdHighlightSelected.Enabled = False
cmdCreateReport.Enabled = False

'Get the GBLakes ID from cboGBLAKES_IDs
Dim varSplit As Variant
Dim i As Long
Dim strGBLAKES_ID As String
Dim lonGBLakesID As Long
lonGBLakesID = 0

If Not IsNull(pGBLakes_WBID_Array(0, 0)) Then
    If cboGBLAKES_IDs.Text <> "" Then
        varSplit = Split(cboGBLAKES_IDs.Text, " - ")
        If varSplit(0) = cboGBLAKES_IDs.Text Then
            MsgBox "You have chosen a code that does not appear in the data.", vbCritical
            Exit Sub
        Else
            strGBLAKES_ID = varSplit(0)
            lonGBLakesID = ReturnWFD_WB_ID(CLng(strGBLAKES_ID))
            If lonGBLakesID <> 0 Then   'if no match then it returns a 0
                cboWBID.Text = lonGBLakesID & " - " & varSplit(1)
            Else
                cboWBID.Text = " - "
            End If
        End If
        Else
            Exit Sub
    End If
End If
End Sub
Private Sub cboWBID_Click()
'get the GBLAKES_ID from the GBLakes ID and send it to cboGBLAKES_IDs.text

Dim varSplit As Variant
Dim i As Long
Dim strGBLakesID As String
Dim lonGBLAKES_ID As Long
lonGBLAKES_ID = 0

If cboWBID.Text <> "" And cboWBID.Text <> " - " Then
    varSplit = Split(cboWBID.Text, " - ")
    If varSplit(0) = cboWBID.Text Then
        MsgBox "You have chosen a code that does not appear in the data.", vbCritical
        Exit Sub
    Else
        strGBLakesID = varSplit(0)
        lonGBLAKES_ID = ReturnGBLakes_ID(CLng(strGBLakesID))
        If lonGBLAKES_ID <> 0 Then   'if no match then it returns a 0
            cboGBLAKES_IDs.Text = lonGBLAKES_ID & " - " & varSplit(1)
        Else
            cboGBLAKES_IDs.Text = " - "
        End If
    End If
    Else
        Exit Sub
End If
cmdCreateReport.Enabled = False
End Sub
Private Sub cboWBID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'get the GBLAKES_ID from the GBLakes ID and send it to cboGBLAKES_IDs.text

Dim varSplit As Variant
Dim i As Long
Dim strGBLakesID As String
Dim lonGBLAKES_ID As Long
lonGBLAKES_ID = 0

If cboWBID.Text <> "" And cboWBID.Text <> " - " Then
    varSplit = Split(cboWBID.Text, " - ")
    If varSplit(0) = cboWBID.Text Then
        MsgBox "You have chosen a code that does not appear in the data.", vbCritical
        Exit Sub
    Else
        strGBLakesID = varSplit(0)
        lonGBLAKES_ID = ReturnGBLakes_ID(CLng(strGBLakesID))
        If lonGBLAKES_ID <> 0 Then   'if no match then it returns a 0
            cboGBLAKES_IDs.Text = lonGBLAKES_ID & " - " & varSplit(1)
        Else
            cboGBLAKES_IDs.Text = " - "
        End If
    End If
    Else
        Exit Sub
End If
cmdCreateReport.Enabled = False
End Sub

Private Sub chkAddPointSource_Click()
    If chkAddPointSource Then
        chkRemoveSelectedPointSources = False
    End If
End Sub
Private Sub chkChangePforNetwork_Click()
If chkChangePforNetwork.Value = True Then
    cmdModCatchmentInputs.Caption = "Modify inputs" & vbCrLf & "(for " & lonChosenGBLAKES_ID & ") and P due to cover for network"
Else
    cmdModCatchmentInputs.Caption = "Modify inputs" & vbCrLf & "(for " & lonChosenGBLAKES_ID & ")"
End If
End Sub

Private Sub chkLoadPointSources_Click()
    If chkLoadPointSources.Value = True Then
        blnLoadNonScenarioPointSources = True
        MsgBox "You must now reload the data to apply this setting", vbCritical
    Else
        blnLoadNonScenarioPointSources = False
        MsgBox "You must now reload the data to apply this setting", vbCritical
    End If

End Sub

Private Sub chkPerCapitaTPLoadRural_Click()
    If chkPerCapitaTPLoadRural.Value = False Then
        chkPerCapitaTPLoadRuralAll.Value = False
    End If

End Sub

Private Sub chkPerCapitaTPLoadRuralAll_Click()
    If chkPerCapitaTPLoadRuralAll.Value = False Then
        chkPerCapitaTPLoadRural.Value = False
        Else
        chkPerCapitaTPLoadRural.Value = True
        If txtPerCapitaTPLoadRural.Text = "" Then
            txtPerCapitaTPLoadRural.Text = 0
        End If
    End If
    
End Sub


Private Sub chkPerCapitaTPLoadUrban_Click()
    If chkPerCapitaTPLoadUrban.Value = False Then
        chkPerCapitaTPLoadUrbanAll.Value = False
    End If

End Sub

Private Sub chkPerCapitaTPLoadUrbanAll_Click()
If chkPerCapitaTPLoadUrbanAll.Value = False Then
    chkPerCapitaTPLoadUrban.Value = False
    Else
    chkPerCapitaTPLoadUrban.Value = True
    If txtPerCapitaTPLoadUrban.Text = "" Then
        txtPerCapitaTPLoadUrban = 0
    End If
End If

End Sub

Private Sub chkRemoveSelectedPointSources_Click()
If chkRemoveSelectedPointSources Then
    chkAddPointSource = False
    chkPerCapitaTPLoadUrban = False
    chkPerCapitaTPLoadRural = False
    chkPerCapitaTPLoadRuralAll.Value = False
    chkPerCapitaTPLoadRuralAll.Value = False

    chkUrbanPop = False
    chkRuralPop = False
    cboPointSourceType.Text = ""
    txtPointSourceAmount.Text = ""
End If
End Sub
Private Sub chkScenarioID_Click()
'fix the scenarioID
If chkScenarioID Then
    FilterScenarioSelection
Else
    Run_MasterOrScenario
End If
End Sub
Private Sub chkScenarioOwner_Click()
If chkScenarioOwner Then
    FilterScenarioSelection
Else
    Run_MasterOrScenario
End If
End Sub
Private Sub chkScenarioName_Click()
If chkScenarioName Then
    FilterScenarioSelection
Else
    Run_MasterOrScenario
End If
End Sub
Private Sub chkScenarioDate_Click()
'fix the date
If chkScenarioDate Then
    FilterScenarioSelection
Else
    Run_MasterOrScenario
End If
End Sub
Private Sub chkScenarioComment_Click()
If chkScenarioComment Then
    FilterScenarioSelection
Else
    Run_MasterOrScenario
End If
End Sub
Sub FilterScenarioSelection()
'use the status of the:
'chkScenarioID, chkScenarioName, chkScenarioOwner, chkScenarioDate, chkScenarioComment to determine what goes in
'cboScenarioID, cboFilterScenarioName, cboFilterScenarioOwner, cboFilterScenarioDate, cboFilterScenarioComment
Dim i As Integer
Dim j As Integer
Dim pCursor As ICursor
Dim pQueryFilt As IQueryFilter2
Dim pRow As IRow
Dim varDateText As Variant
Dim blnFirstFound As Boolean
Dim blnSecondFound As Boolean
Dim blnThirdFound As Boolean
Dim blnFourthFound As Boolean
Dim blnFifthFound As Boolean
Dim blnScenarioIDUsed As Boolean
Dim blnScenarioNameUsed As Boolean
Dim blnScenarioOwnerUsed As Boolean
Dim blnScenarioDateUsed As Boolean
Dim blnScenarioCommentUsed As Boolean

tglMasterOrScenario.Caption = "Click to load from master data"

'get the number of chkScenario selected - this is needed to create the query filter
j = 0
If chkScenarioID Then
    j = j + 1
End If
If chkScenarioName Then
    j = j + 1
End If
If chkScenarioOwner Then
    j = j + 1
End If
If chkScenarioDate Then
    j = j + 1
End If
If chkScenarioComment Then
    j = j + 1
End If

'initialise the pull down menus
For i = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(i).Name = strScenarioTbl Then
        Set pScenario = pTabColl.StandaloneTable(i)
    End If
Next
intScenarioIDField = pScenario.Table.FindField("ScenarioID")
intScenarioNameField = pScenario.Table.FindField("ScenarioName")
intScenarioCreatorField = pScenario.Table.FindField("ScenarioCreator")
intScenarioCreationDateField = pScenario.Table.FindField("ScenarioCreationDate")
intScenarioCommentField = pScenario.Table.FindField("Comment")
intScenarioRegionField = pScenario.Table.FindField("Region")
If intScenarioIDField = -1 Then
    MsgBox "Cannot find field ScenarioID", vbCritical
    Exit Sub
End If

'#######################################################################################
'Implement user selection
'#######################################################################################
If j = 1 Then
    Set pQueryFilt = New QueryFilter
    If chkScenarioID Then
        pQueryFilt.WhereClause = "ScenarioID = " & cboScenarioID.Value
    End If
    If chkScenarioName Then
        pQueryFilt.WhereClause = "ScenarioName = '" & cboFilterScenarioName.Value & "'"
    End If
    If chkScenarioOwner Then
        pQueryFilt.WhereClause = "ScenarioCreator = '" & cboFilterScenarioOwner.Value & "'"
    End If
    If chkScenarioDate Then
    'date is stored in the format: #11-30-2009 00:00:00#
        varDateText = Split(cboFilterScenarioDate.Value, "/")
        pQueryFilt.WhereClause = "ScenarioCreationDate = #" & varDateText(1) & "-" & varDateText(0) & "-" & varDateText(2) & " 00:00:00#"
    End If
    If chkScenarioComment Then
        pQueryFilt.WhereClause = "Comment = '" & cboFilterScenarioComment.Value & "'"
    End If
    cboScenarioID.Clear
    cboFilterScenarioName.Clear
    cboFilterScenarioOwner.Clear
    cboFilterScenarioDate.Clear
    cboFilterScenarioComment.Clear
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    While Not pRow Is Nothing
        cboScenarioID.AddItem pRow.Value(intScenarioIDField)
        cboFilterScenarioName.AddItem pRow.Value(intScenarioNameField)
        cboFilterScenarioOwner.AddItem pRow.Value(intScenarioCreatorField)
        cboFilterScenarioDate.AddItem Day(pRow.Value(intScenarioCreationDateField)) & "/" & _
        Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
        cboFilterScenarioComment.AddItem pRow.Value(intScenarioCommentField)
        Set pRow = pCursor.NextRow
    Wend
    'once again so that the value in the cboboxes is the first found
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    cboScenarioID.Value = pRow.Value(intScenarioIDField)
    cboFilterScenarioName.Value = pRow.Value(intScenarioNameField)
    cboFilterScenarioOwner.Value = pRow.Value(intScenarioCreatorField)
    cboFilterScenarioDate.Value = Day(pRow.Value(intScenarioCreationDateField)) & "/" & _
    Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
    cboFilterScenarioComment.Value = pRow.Value(intScenarioCommentField)
Else
    Set pQueryFilt = New QueryFilter
    blnScenarioIDUsed = False
    blnScenarioNameUsed = False
    blnScenarioOwnerUsed = False
    blnScenarioDateUsed = False
    blnScenarioCommentUsed = False
    'get the first part of the query
    blnFirstFound = False
    blnSecondFound = False
    blnThirdFound = False
    blnFourthFound = False
    If chkScenarioID Then
        pQueryFilt.WhereClause = "ScenarioID = " & cboScenarioID.Value & " and "
        blnFirstFound = True
        blnScenarioIDUsed = True
    End If
    If chkScenarioName And Not blnFirstFound Then
        pQueryFilt.WhereClause = "ScenarioName = '" & cboFilterScenarioName.Value & "'" & " and "
        blnFirstFound = True
        blnScenarioNameUsed = True
    End If
    If chkScenarioOwner And Not blnFirstFound Then
        pQueryFilt.WhereClause = "ScenarioCreator = '" & cboFilterScenarioOwner.Value & "'" & " and "
        blnFirstFound = True
        blnScenarioOwnerUsed = True
    End If
    If chkScenarioDate And Not blnFirstFound Then
        varDateText = Split(cboFilterScenarioDate.Value, "/")
        pQueryFilt.WhereClause = "ScenarioCreationDate = #" & varDateText(1) & "-" & varDateText(0) & "-" & varDateText(2) & " 00:00:00#" & " and "
        blnFirstFound = True
        blnScenarioDateUsed = True
    End If

'get the second part of the query
    If chkScenarioName And Not blnSecondFound And Not blnScenarioNameUsed Then
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioName = '" & cboFilterScenarioName.Value & "'"
        If j > 2 Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
        End If
        blnSecondFound = True
        blnScenarioNameUsed = True
    End If
    If chkScenarioOwner And Not blnSecondFound And Not blnScenarioOwnerUsed Then
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioCreator = '" & cboFilterScenarioOwner.Value & "'"
        If j > 2 Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
        End If
        blnSecondFound = True
        blnScenarioOwnerUsed = True
    End If
    If chkScenarioDate And Not blnSecondFound And Not blnScenarioDateUsed Then
        varDateText = Split(cboFilterScenarioDate.Value, "/")
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioCreationDate = #" & varDateText(1) & "-" & varDateText(0) & "-" & varDateText(2) & " 00:00:00#"
        If j > 2 Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
        End If
        blnSecondFound = True
        blnScenarioDateUsed = True
    End If
    If chkScenarioComment And Not blnSecondFound And Not blnScenarioCommentUsed Then
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & "Comment = '" & cboFilterScenarioComment.Value & "'"
        If j > 2 Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
        End If
        blnSecondFound = True
        blnScenarioCommentUsed = True
    End If
    
    If j > 2 Then
'get the third part of the query
        If chkScenarioOwner And Not blnScenarioOwnerUsed Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioCreator = '" & cboFilterScenarioOwner.Value & "'"
            If j > 3 Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
            End If
            blnThirdFound = True
            blnScenarioOwnerUsed = True
        End If
        If chkScenarioDate And Not blnThirdFound And Not blnScenarioDateUsed Then
            varDateText = Split(cboFilterScenarioDate.Value, "/")
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioCreationDate = #" & varDateText(1) & "-" & varDateText(0) & "-" & varDateText(2) & " 00:00:00#"
            If j > 3 Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
            End If
            blnThirdFound = True
            blnScenarioDateUsed = True
        End If
        If chkScenarioComment And Not blnThirdFound And Not blnScenarioCommentUsed Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "Comment = '" & cboFilterScenarioComment.Value & "'"
            If j > 3 Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
            End If
            blnThirdFound = True
            blnScenarioCommentUsed = True
        End If
    End If

    If j > 3 Then
'get the fourth part of the query
        If chkScenarioDate And Not blnFourthFound And Not blnScenarioDateUsed Then
            varDateText = Split(cboFilterScenarioDate.Value, "/")
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "ScenarioCreationDate = #" & varDateText(1) & "-" & varDateText(0) & "-" & varDateText(2) & " 00:00:00#"
            If j > 3 Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
            End If
            blnFourthFound = True
            blnScenarioDateUsed = True
        End If
        If chkScenarioComment And Not blnFourthFound And Not blnScenarioCommentUsed Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "Comment = '" & cboFilterScenarioComment.Value & "'"
            If j > 3 Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " and "
            End If
            blnFourthFound = True
            blnScenarioCommentUsed = True
        End If
    End If

    If j > 4 Then
'get the fifth part of the query
        If chkScenarioComment And Not blnFifthFound And Not blnScenarioCommentUsed Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & "Comment = '" & cboFilterScenarioComment.Value & "'"
            blnFifthFound = True
            blnScenarioCommentUsed = True
        End If
    End If
    
    cboScenarioID.Clear
    cboFilterScenarioName.Clear
    cboFilterScenarioOwner.Clear
    cboFilterScenarioDate.Clear
    cboFilterScenarioComment.Clear
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    While Not pRow Is Nothing
        cboScenarioID.AddItem pRow.Value(intScenarioIDField)
        cboFilterScenarioName.AddItem pRow.Value(intScenarioNameField)
        cboFilterScenarioOwner.AddItem pRow.Value(intScenarioCreatorField)
        cboFilterScenarioDate.AddItem Day(pRow.Value(intScenarioCreationDateField)) & "/" & Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
        cboFilterScenarioComment.AddItem pRow.Value(intScenarioCommentField)
        Set pRow = pCursor.NextRow
    Wend
    'once again so that the value in the cboboxes is the first found
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    cboScenarioID.Value = pRow.Value(intScenarioIDField)
    cboFilterScenarioName.Value = pRow.Value(intScenarioNameField)
    cboFilterScenarioOwner.Value = pRow.Value(intScenarioCreatorField)
    cboFilterScenarioDate.Value = Day(pRow.Value(intScenarioCreationDateField)) & "/" & Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
    cboFilterScenarioComment.Value = pRow.Value(intScenarioCommentField)
End If

End Sub
Private Sub cmdAbout_Click()
MsgBox "Please consult the User Manual for detailed instructions on the use of this application. " & vbCrLf & _
        "This application was produced by the James Hutton Institute, Aberdeen." & vbCrLf & _
        "PLUS+ Copyright (c) 2012 James Hutton Institute" & vbCrLf & "This program comes with ABSOUTELY NO WARRANTY" & vbCrLf & _
        "This is free software, and you are welcome to redistribute it" & vbCrLf & "under the terms of the GNU General Public License", vbInformation, "Information"
End Sub
Private Sub cmdCalcTP_Click()
'the point source should be the size of the loaded scenario if appropriate
'this initialises the array to have one line per site within the network
'as this is always done before the addition of any point sources in scenario creation it is always available.

CalculateTP
cmdCreateReport.Enabled = True
'report to lstViewSupplement
lstViewSupplement.ColumnHeaders.Clear
Dim intArrayColumnWidths(7) As Integer
intArrayColumnWidths(1) = 54
intArrayColumnWidths(2) = 109
intArrayColumnWidths(3) = 152
intArrayColumnWidths(4) = 78
intArrayColumnWidths(5) = 87
intArrayColumnWidths(6) = 69
intArrayColumnWidths(7) = 90
Dim strArrayColumnHeadings(7) As String
strArrayColumnHeadings(1) = "GB Lakes ID"
strArrayColumnHeadings(2) = "Local run-off (m" & Chr(179) & " per year)"
strArrayColumnHeadings(3) = "Local & up-stream run-off (m" & Chr(179) & " per year)"
strArrayColumnHeadings(4) = "OECD numerator"
strArrayColumnHeadings(5) = "OECD exponent den."
strArrayColumnHeadings(6) = "Mean depth (m)"
strArrayColumnHeadings(7) = "Loch area (m" & Chr(178) & ")"

lstViewSupplement.ListItems.Clear

Dim i As Integer
For i = 1 To 7
   Set SupplementColumnHeaders = lstViewSupplement.ColumnHeaders.Add()
   SupplementColumnHeaders.Text = strArrayColumnHeadings(i)
   SupplementColumnHeaders.Width = intArrayColumnWidths(i)
Next

lstViewSupplement.ListItems.Clear

For i = 0 To UBound(CatchNetRship, 1)
    Set List_ItemSupplement = lstViewSupplement.ListItems.Add
    List_ItemSupplement = CatchNetRship(i, 0)
    List_ItemSupplement.SubItems(1) = Format(CatchNetRship(i, 2), "#,#.0")
    List_ItemSupplement.SubItems(2) = Format(CatchNetRship(i, 3), "#,#.0")
    List_ItemSupplement.SubItems(3) = CatchNetRship(i, 7)
    List_ItemSupplement.SubItems(4) = CatchNetRship(i, 8)
    List_ItemSupplement.SubItems(5) = Format(CatchNetRship(i, 23), "#,#.00")
    If CatchNetRship(i, 24) = 0 Then
        List_ItemSupplement.SubItems(6) = "0"
    Else
        List_ItemSupplement.SubItems(6) = Format(CatchNetRship(i, 24) * 10000, "#,#")
    End If
Next

End Sub
Sub CalculateTP()
'############################################################################
'Calculate Total P - this is one of the key routines
'############################################################################

'test the MSComctlLib library is loaded
'If Not CheckReferencesAttached Then
'    MsgBox "Warning, the Microsoft ImageList Control 6.0 library is not referenced in the Visual Basic environment." _
'    & vbCrLf & "This must be referenced for the tool to function. Please consult the user guide.", vbCritical
'    Exit Sub
'End If

Dim intTempCounter As Integer
Dim H As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim iOIDList() As Long
Dim lonRowScroller As Long

Dim pCursor As ICursor
Dim pRow As IRow
Dim dblCumulativeLocalRunoff As Double
Dim intGBLAKES_IDToFind As Long
Dim blnWhileMaxGBLAKES_IDNotFound As Boolean
Dim lonWorkingGBLAKES_ID As Long
Dim intWhileLoopCounter As Long
Dim column_header As columnHeader   'for lvwCatchmentInfo

intIndexSelectedGBLAKES_ID = 99999
For i = 0 To UBound(CatchNetRship, 1)
    If CatchNetRship(i, 0) = lonChosenGBLAKES_ID Then
            intIndexSelectedGBLAKES_ID = i
    End If
Next
If intIndexSelectedGBLAKES_ID = 99999 Then
    MsgBox "Cannot find the index of the selected GBLAKES_ID.", vbCritical
    Exit Sub
End If

'###################################################################################################################################
'For the selected catchment calculate the TP, this involves starting at the 0 order water bodies and calculating their TP and cascading down,
'calculating the TP for the whole network
'This section uses the following tables:
'   tblLoadPrecursor for LocalRunoff, OECDDenominator, OECDExponentDenominator, LochArea, LochMeanDepth, LocalArea
'   tblExports for processing later - it's the Land Cover of Scotland (LCS) look up table
'   tblCatchP - P for each of the land cover classes in each catchment it has a P value - used with tblExports to calculate scenarios
'   strTblLoadPrecursorName strTblCatchPName strTblExportsName
'###################################################################################################################################

'initialise the CatchNetRship()
For i = 0 To UBound(CatchNetRship, 1)
    CatchNetRship(i, 3) = 0
    CatchNetRship(i, 4) = 0
    CatchNetRship(i, 9) = 0
    CatchNetRship(i, 10) = 0
    CatchNetRship(i, 11) = 0
    CatchNetRship(i, 12) = 0
    CatchNetRship(i, 14) = 0
    CatchNetRship(i, 15) = 0
    CatchNetRship(i, 16) = 0
    CatchNetRship(i, 17) = 0
    CatchNetRship(i, 18) = 0
    CatchNetRship(i, 19) = 0
    CatchNetRship(i, 20) = ""
    CatchNetRship(i, 21) = ""
    CatchNetRship(i, 22) = 0
    CatchNetRship(i, 23) = 0
    CatchNetRship(i, 24) = 0
    CatchNetRship(i, 25) = 0
    CatchNetRship(i, 26) = 0
Next

'use the following arrays created in an earlier routine
'   CatchNetRship(x,0) = GBLAKES_ID, CatchNetRship(x,1) = "Separate Branch", "Chosen GBLAKES_ID", "Is Downstream from", "Is Upstream of"
'   arrayFlowRouting(x,y) - arrayFlowRouting(x,0) = upstream, arrayFlowRouting(x, 1) - downstream
'   lonOrderMatchArray(i)

'start with order 0 as there is no upstream runoff to include
'runoff = (LocalRunoff * LocalCatchmentArea) + Sum(UpstreamRunoff)
'       = (tblLoadPrecursor.LocalRunoff * tblLoadPrecursor.LocalArea) + Sum(UpstreamRunoff)

'calculate the Runoff R = jAL + sum(Upstream runoff)
'                        j= local runoff in m/yr, AL is the local catchment area in m2

GetFieldIndices 'find the required fields for the tables

'calculate s - the sewage load for each catchment
Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
pQueryFilt.SubFields = "GBLAKES_ID,Urb_Rur,Load,Pop"
'pQueryFilt.SubFields = "GBLAKES_ID,Urb_Rur,NoUrb_Load,NoUrb_Pop"

'loop to create the WhereClause
pQueryFilt.WhereClause = "GBLAKES_ID = "
For i = 0 To UBound(CatchNetRship, 1)
    pQueryFilt.WhereClause = pQueryFilt.WhereClause & CatchNetRship(i, 0)
    If i < UBound(CatchNetRship, 1) Then
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & " or GBLAKES_ID = "
    End If
Next

'for sewage there is one record per catchment - GBLAKES_ID, Urban load, rural load, Urban Pop, Rural Pop
ReDim varCatchmentSewage(UBound(CatchNetRship, 1), 4)
'for point-source one record per source varPointSource() - GBLAKES_ID, source type, source amount, scenario id
For i = 0 To UBound(varCatchmentSewage, 1)
    varCatchmentSewage(i, 0) = 0
    varCatchmentSewage(i, 1) = 0
    varCatchmentSewage(i, 2) = 0
    varCatchmentSewage(i, 3) = 0
    varCatchmentSewage(i, 4) = 0
Next
Dim blnArrayIsNotEmpty As Boolean
'populate first position with the GBLAKES_IDs
For i = 0 To UBound(CatchNetRship, 1)
    varCatchmentSewage(i, 0) = CatchNetRship(i, 0)
Next

blnArrayIsNotEmpty = False
If blnScenarioLoaded Then
    'get the data from the Scenario point source table
    'and must also ensure that the user entered values (post-scenario loading) are included
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
    Set pCursor = pCatchmentSewageTable_S.Table.Search(pQueryFilt, False)
    intSewageGBLAKES_IDField = pCatchmentSewageTable_S.Table.FindField("GBLAKES_ID")
    intUrb_RurField = pCatchmentSewageTable_S.Table.FindField("Urb_Rur")
    intLoadField = pCatchmentSewageTable_S.Table.FindField("Load")
    intPopulationField = pCatchmentSewageTable_S.Table.FindField("Pop")
    
'the 'Get catchment info.' button will initiate the load of the scenario point source
'subsequent additions and removals will be handled by 'Apply changes'
    If blnModifyOtherPointSourceLoad = True Then
        If IsNumeric(txtPointSourceAmount.Text) Then
            For i = 0 To UBound(varPointSource, 1)
                If varPointSource(i, 0) = lonChosenGBLAKES_ID Then
                    If varPointSource(i, 2) = 0 Then
                        varPointSource(i, 2) = CDbl(txtPointSourceAmount.Text)
                        varPointSource(i, 1) = cboPointSourceType.Value
                    End If
                End If
            Next
        End If
    Else
    End If
Else

    Set pCursor = pCatchmentSewageTable.Table.Search(pQueryFilt, False)
End If
'########################################################################################################
'The population model gives a continuous surface of population density, so a catchment can have 0.134 people
'The load in CatchmentSewage is dependent on this precise value, but I report the population rounded to a whole person
'This could lead to problems where a scenario with the same number of people has a different output
'A boolean variable is placed here to change processing to use a calculated sewage load dependent on the rounded number of people,
'but retaining the option to switch it off should it be required at a later date.
'########################################################################################################
Dim blnUseWholePeople As Boolean
blnUseWholePeople = True 'This is a switch to allow the use of real or partial people, change it to false to use non-whole numbers

DisplaySewageInfo
Set pRow = pCursor.NextRow
While Not pRow Is Nothing
    For i = 0 To UBound(varCatchmentSewage, 1)
        If varCatchmentSewage(i, 0) = pRow.Value(intSewageGBLAKES_IDField) Then
            If pRow.Value(intUrb_RurField) = "Urban" Then
                'Debug.Print "Values: " & pRow.Value(0) & " - " & pRow.Value(1) & " - " & pRow.Value(2) & " - " & pRow.Value(3) & " - " & pRow.Value(4) & " - " & pRow.Value(5) & " - " & pRow.Value(6) ' & " - " & pRow.Value(7)
                varCatchmentSewage(i, 1) = pRow.Value(intLoadField)
                varCatchmentSewage(i, 3) = pRow.Value(intPopulationField)
                If blnUseWholePeople Then
                    varCatchmentSewage(i, 1) = CLng(varCatchmentSewage(i, 3)) * dblUrbanPerCapitaTPLoad 'dblRuralPerCapitaTPLoad
                End If
            End If
            If pRow.Value(intUrb_RurField) = "Rural" Then
                varCatchmentSewage(i, 2) = pRow.Value(intLoadField)
                varCatchmentSewage(i, 4) = pRow.Value(intPopulationField)
                If blnUseWholePeople Then
                    varCatchmentSewage(i, 2) = CInt(varCatchmentSewage(i, 4)) * dblRuralPerCapitaTPLoad
                End If
            End If
            
        End If
    Next
    Set pRow = pCursor.NextRow
Wend

'will retain varCatchmentSewage() as the read-in data, in scenarios we may change CatchNetRship()
'populate CatchNetRship(i,16, 17 & 18) with the sewage and population from varCatchmentSewage
For i = 0 To UBound(CatchNetRship, 1)
    CatchNetRship(i, 16) = 0    'Urban Load
    CatchNetRship(i, 17) = 0    'Rural Load
    CatchNetRship(i, 18) = 0    'Population from CatchmentSewage table
    CatchNetRship(i, 19) = 0    'Population from CatchmentSewage table
    CatchNetRship(i, 22) = 0    'Amount of user entered point source
Next

blnArrayIsNotEmpty = False
On Error Resume Next
blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
For i = 0 To UBound(CatchNetRship, 1)
    For j = 0 To UBound(varCatchmentSewage, 1)
        If CatchNetRship(i, 0) = varCatchmentSewage(j, 0) Then 'test on the sitecode
            CatchNetRship(i, 16) = varCatchmentSewage(j, 1)
            CatchNetRship(i, 17) = varCatchmentSewage(j, 2)
            CatchNetRship(i, 18) = varCatchmentSewage(j, 3)
            CatchNetRship(i, 19) = varCatchmentSewage(j, 4)
            Exit For
        End If
    Next
    If blnArrayIsNotEmpty Then
        For j = 0 To UBound(varPointSource, 1)  'add the point source
        'this section ensures that a scenario with point sources from multiple sites are correctly apportioned
            If CatchNetRship(i, 0) = varPointSource(j, 0) Then
                CatchNetRship(i, 22) = CatchNetRship(i, 22) + varPointSource(j, 2)
            End If
        Next
    Else 'there are no point sources in the scenario
        CatchNetRship(i, 22) = 0
    End If
Next

DisplaySewageInfo   'note this is in twice because the subroutine produces a value needed elsewhere
'#######################################################################################
'Implement user modifications - these apply to baseline and to loaded scenario data
'#######################################################################################
If blnModifySewageLoad Then
'depending on the check boxes in Frame_Modify_Load make changes
    If chkPerCapitaTPLoadUrban Then 'if changing per capita TP load for urban
    'must do this in two stages in case the population is also being changed
        If IsNumeric(txtPerCapitaTPLoadUrban.Text) Then 'check that a number has been entered
            If arrayPerCapitaTPLoads(0, 0) = "Urban" Then
                If chkUrbanPop Then 'if changing urban population too
                    If Not IsNumeric(txtUrbanPop.Text) Then
                            MsgBox "The text in your Urban population box must be a number, exiting...", vbCritical
                            Exit Sub
                    End If
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = CLng(txtUrbanPop.Text) '18 is the urban population, 16 is the urban load
                    If chkPerCapitaTPLoadUrbanAll Then
                    'update all, overwriting the info from DisplaySewageInfo
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                            'Debug.Print "Loch " & CatchNetRship(i, 0) & " has a load of " & CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                        Next
                    Else
                    'only selected
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 18)
                    End If
                Else
                    If chkPerCapitaTPLoadUrbanAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 18)
                    End If
                End If
            ElseIf arrayPerCapitaTPLoads(1, 0) = "Urban" Then
                If chkUrbanPop Then 'if changing urban population too
                    If Not IsNumeric(txtUrbanPop.Text) Then
                            MsgBox "The text in your Urban population box must be a number, exiting...", vbCritical
                            Exit Sub
                    End If
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = CLng(txtUrbanPop.Text) '18 is the urban population, 16 is the urban load
                    If chkPerCapitaTPLoadUrbanAll Then
                    'update all, overwriting the info from DisplaySewageInfo
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                            'Debug.Print "Loch " & CatchNetRship(i, 0) & " has a load of " & CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                        Next
                    Else
                    'only selected
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 18)
                    End If
                Else
                    If chkPerCapitaTPLoadUrbanAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(i, 18)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CDbl(txtPerCapitaTPLoadUrban.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 18)
                    End If
                End If
            End If
        Else
            MsgBox "The text in your Per Capita TP Load Urban box must be a number, exiting...", vbCritical
            Exit Sub
        End If
    End If
    If chkPerCapitaTPLoadRural Then
    'must do this in two stages in case the population is also being changed
        If IsNumeric(txtPerCapitaTPLoadRural.Text) Then
            If arrayPerCapitaTPLoads(0, 0) = "Rural" Then
                If chkRuralPop Then
                    If Not IsNumeric(txtRuralPop.Text) Then
                            MsgBox "The text in your Rural population box must be a number, exiting...", vbCritical
                            Exit Sub
                    End If
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = CLng(txtRuralPop.Text)
                    
                    If chkPerCapitaTPLoadRuralAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(i, 19)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19) 'was: arrayPerCapitaTPLoads(0, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                    End If
                Else
                    If chkPerCapitaTPLoadRuralAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(i, 19)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19) 'was: arrayPerCapitaTPLoads(0, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                    End If
                End If
            ElseIf arrayPerCapitaTPLoads(1, 0) = "Rural" Then
                If chkRuralPop Then
                    If Not IsNumeric(txtRuralPop.Text) Then
                            MsgBox "The text in your Rural population box must be a number, exiting...", vbCritical
                            Exit Sub
                    End If
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = CLng(txtRuralPop.Text)
                    If chkPerCapitaTPLoadRuralAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(i, 19)  'was: arrayPerCapitaTPLoads(1, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)  'was: arrayPerCapitaTPLoads(1, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                    End If
                Else
                    If chkPerCapitaTPLoadRuralAll Then
                        For i = 0 To UBound(CatchNetRship, 1)
                            CatchNetRship(i, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(i, 19)  'was: arrayPerCapitaTPLoads(1, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                        Next
                    Else
                        CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CDbl(txtPerCapitaTPLoadRural.Text) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)  'was: arrayPerCapitaTPLoads(1, 1) * CatchNetRship(intIndexSelectedGBLAKES_ID, 19)
                    End If
                End If
            End If
        Else
            MsgBox "The text in your Per Capita TP Load Rural box must be a number, exiting...", vbCritical
            Exit Sub
        End If
    End If
'do the pop - this doesn't need the global switch - this is goverened by whether or not chkPerCapitaTPLoadUrban is checked
    If chkUrbanPop Then
        If IsNumeric(txtUrbanPop.Text) Then
            CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = CLng(txtUrbanPop.Text)
            If chkPerCapitaTPLoadUrban Then
                If Not IsNumeric(txtPerCapitaTPLoadUrban.Text) Then
                    MsgBox "The text in your Urban Load box must be a number, exiting...", vbCritical
                    Exit Sub
                End If
                CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CLng(txtUrbanPop.Text) * CDbl(txtPerCapitaTPLoadUrban.Text)
            Else
                'get the load from the matrix (unless we want to use the global setting)
                If arrayPerCapitaTPLoads(0, 0) = "Urban" Then
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CLng(txtUrbanPop.Text) * arrayPerCapitaTPLoads(0, 1)
                Else
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 16) = CLng(txtUrbanPop.Text) * arrayPerCapitaTPLoads(1, 1)
                End If
            End If
        Else
            MsgBox "The text in your Urban Population box must be a number, exiting...", vbCritical
            Exit Sub
        End If
    End If
    If chkRuralPop Then
        If IsNumeric(txtRuralPop.Text) Then
            CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = CLng(txtRuralPop.Text)
            If chkPerCapitaTPLoadRural Then
                If Not IsNumeric(txtPerCapitaTPLoadRural.Text) Then
                    MsgBox "The text in your Rural Load box must be a number, exiting...", vbCritical
                    Exit Sub
                End If
                CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CLng(txtRuralPop.Text) * CDbl(txtPerCapitaTPLoadRural.Text)
            Else
                'get the load from the matrix
                If arrayPerCapitaTPLoads(0, 0) = "Rural" Then
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CLng(txtRuralPop.Text) * arrayPerCapitaTPLoads(0, 1)
                Else
                    CatchNetRship(intIndexSelectedGBLAKES_ID, 17) = CLng(txtRuralPop.Text) * arrayPerCapitaTPLoads(1, 1)
                End If
            End If
        Else
            MsgBox "The text in your Rural Population box must be a number, exiting...", vbCritical
            Exit Sub
        End If
    End If
End If
'#######################################################################################
'Calculate jAL (local run off * area) for all catchments in the network from pLoadPrecursorTable
'#######################################################################################
If blnScenarioLoaded Then
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
    Set pCursor = pLoadPrecursorTable_S.Table.Search(pQueryFilt, False)
    intPrecursorGBLAKES_IDField = pLoadPrecursorTable_S.Table.FindField("GBLAKES_ID")
    intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff")
    'intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff_M10")
    'intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff_M5")
    'intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff_P5")
    'intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff_P10")
    intLocalAreaField = pLoadPrecursorTable_S.Table.FindField("LocalArea")
Else
    ReDim iOIDList(lonNumGBLAKES_IDs)
    For lonRowScroller = 0 To lonNumGBLAKES_IDs 'need to be careful with this - this is the row index - if the table doesn't start with 0 then the
                                                'last record could be missed - this will attempt the last record for a table starting with 1
        iOIDList(lonRowScroller) = lonRowScroller
    Next
    Set pCursor = pLoadPrecursorTable.Table.GetRows(iOIDList, True)
End If
Set pRow = pCursor.NextRow
While Not pRow Is Nothing
    For i = 0 To UBound(CatchNetRship, 1)
        If CatchNetRship(i, 0) = pRow.Value(intPrecursorGBLAKES_IDField) Then
            CatchNetRship(i, 2) = pRow.Value(intLocalRunoffField) * pRow.Value(intLocalAreaField)
            If IsNull(CatchNetRship(i, 2)) Then
                'MsgBox "Site " & CatchNetRship(i, 0) & " contains NULL values for local runoff or local area and so cannot be processed.", vbCritical
                Exit Sub
            End If
            If blnScenarioLoaded Then
                'LochArea is in hectares
                CatchNetRship(i, 5) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("LochArea")) * pRow.Value(pLoadPrecursorTable_S.Table.FindField("LochMeanDepth")) * 10000   'calculate loch volume
                CatchNetRship(i, 7) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("OECDDenominator"))
                CatchNetRship(i, 8) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("OECDExponentDenominator"))
                CatchNetRship(i, 13) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("LochOrder"))
                CatchNetRship(i, 23) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("LochMeanDepth"))
                CatchNetRship(i, 24) = pRow.Value(pLoadPrecursorTable_S.Table.FindField("LochArea"))
            Else
                'LochArea is in hectares
                CatchNetRship(i, 5) = pRow.Value(intLochAreaField) * pRow.Value(intLochDepthField) * 10000   'calculate loch volume
                Dim useTableOECDVaules As Boolean
                useTableOECDVaules = False
                If (useTableOECDVaules) Then
                    CatchNetRship(i, 7) = pRow.Value(intOECDDenominatorField)
                    CatchNetRship(i, 8) = pRow.Value(intOECDExponentDenominatorField)
                Else
                    Dim dblDepthThreshold As Double
                    dblDepthThreshold = 2.3
                    If (pRow.Value(intLochDepthField) >= dblDepthThreshold) Then
                        CatchNetRship(i, 7) = 1.55
                        CatchNetRship(i, 8) = 0.82
                    Else
                        CatchNetRship(i, 7) = 1.02
                        CatchNetRship(i, 8) = 0.88
                    End If
                End If
                CatchNetRship(i, 13) = pRow.Value(intOrderField)
                CatchNetRship(i, 23) = pRow.Value(intLochDepthField)
                CatchNetRship(i, 24) = pRow.Value(intLochAreaField)
                'Debug.Print "For " & CatchNetRship(i, 0) & " a is: " & CatchNetRship(i, 7) & " and b is " & CatchNetRship(i, 8) & " and the threshold is " & dblDepthThreshold
            End If
        End If
    Next
    Set pRow = pCursor.NextRow
Wend

If intMatchingGBLAKES_IDs = 0 Then
'for stand alone networks "sum of the current and upstream local runoffs" = "jAL, local runoff in m/yr x local catchment area in m2"
    CatchNetRship(0, 3) = CatchNetRship(0, 2)
Else
'find the GBLAKES_ID of the downstream catchment
For i = 0 To UBound(CatchNetRship, 1)
    For j = 0 To UBound(arrayFlowRouting, 1)
        If arrayFlowRouting(j, 0) = CatchNetRship(i, 0) Then
            CatchNetRship(i, 12) = arrayFlowRouting(j, 1)
            Exit For
        End If
    Next
Next
'#######################################################################################
'Calculate the runoff - step through the GBLAKES_IDs, processing only the zero order catchments down to the end
'#######################################################################################
For H = 0 To UBound(lonOrderMatchArray())
    If lonOrderMatchArray(H) = 0 Then
        CatchNetRship(H, 3) = CatchNetRship(H, 2) 'for order zero the sum = the local
        CatchNetRship(H, 4) = 1  'set the "it's processed" flag
        dblCumulativeLocalRunoff = CatchNetRship(H, 2)
        blnWhileMaxGBLAKES_IDNotFound = True
        lonWorkingGBLAKES_ID = CatchNetRship(H, 0)
        intWhileLoopCounter = 0
        While blnWhileMaxGBLAKES_IDNotFound
            'step through arrayFlowRouting(), following any paths to find the relation to lonChosenGBLAKES_ID for each of the catchments in lonGBLAKES_IDNetworkMatchArray()
            For j = 0 To UBound(arrayFlowRouting)
                If arrayFlowRouting(j, 0) = lonWorkingGBLAKES_ID Then
                'always process with arrayFlowRouting(j, 1)
                    'find the location in CatchNetRship(i, 0)
                    For i = 0 To UBound(CatchNetRship, 1)
                        If CatchNetRship(i, 0) = arrayFlowRouting(j, 1) Then
                            If CatchNetRship(i, 4) = 0 Then  'it's not processed
                                dblCumulativeLocalRunoff = dblCumulativeLocalRunoff + CatchNetRship(i, 2)
                                CatchNetRship(i, 3) = dblCumulativeLocalRunoff
                                CatchNetRship(i, 4) = 1
                            Else
                                'cascade down without incrementing the local runoff to the cumulative - that's already been taken account of below this catchment
                                CatchNetRship(i, 3) = CatchNetRship(i, 3) + dblCumulativeLocalRunoff
                            End If
                            If arrayFlowRouting(j, 1) = lonGBLAKES_IDWithMaxOrder Then
                                For k = 0 To UBound(lonGBLAKES_IDNetworkMatchArray, 1)
                                    If CatchNetRship(k, 0) = lonGBLAKES_IDWithMaxOrder Then
                                        blnWhileMaxGBLAKES_IDNotFound = False
                                    End If
                                Next
                            End If
                        End If
                    Next
                    lonWorkingGBLAKES_ID = arrayFlowRouting(j, 1)
                    If lonWorkingGBLAKES_ID = lonGBLAKES_IDWithMaxOrder Then
                        blnWhileMaxGBLAKES_IDNotFound = False
                    End If
                    Exit For
                End If
            Next
            intWhileLoopCounter = intWhileLoopCounter + 1
            If intWhileLoopCounter = 500 Then
            'this is here to trap a Catchment that does not flow into another in case of problems in the source data
            'not encountered during testing so it is here just in case
                MsgBox "The program appears to be stuck in a loop", vbCritical
                Exit Sub
            End If
        Wend
    End If
Next

'Calculate Tw - water residence time
For i = 0 To UBound(CatchNetRship, 1)
    CatchNetRship(i, 6) = CatchNetRship(i, 5) / CatchNetRship(i, 3)
    'Debug.Print "Calculated water residence time for : "; CatchNetRship(i, 0) & " is " & CatchNetRship(i, 6)
Next

End If 'end of If intMatchingGBLAKES_IDs <> 0 Then

For i = 0 To UBound(CatchNetRship, 1)   'calc Tw
    CatchNetRship(i, 6) = CatchNetRship(i, 5) / CatchNetRship(i, 3)
    'Debug.Print "Calculated water residence time for : "; CatchNetRship(i, 0) & " is " & CatchNetRship(i, 6)
Next

'#######################################################################################
'Calculate TP = CatchNetRship(i, 7) * ((tblCatchP.P + s + x + y + (sum upstream inputs))/(1+ root Tw [i.e. residence time]) exp CatchNetRship(i, 8)
'as before calculate from order 0 and cascade down, but only one level at a time
'calculate CatchNetRship(i, 9) = Sum P in tblCatchP for each of the landcover types for each GBLAKES_ID
'#######################################################################################
'Read the catchP table into dblCatchP() array and then refine it to sum per site code in CatchNetRship(i, 9)
'filter for lonChosenNetwork, take account of master source or scenario data
Set pQueryFilt = New QueryFilter
If blnScenarioLoaded Then
    pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
    lontblCatchPRecords = pCatchPTable_S.Table.RowCount(pQueryFilt)
    ReDim dblCatchP(lontblCatchPRecords - 1, 6)
    Set pCursor = pCatchPTable_S.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    lonRowScroller = 0
    
    While Not pRow Is Nothing
    'read P from each of the different landcover types for each catchment in the network
        dblCatchP(lonRowScroller, 0) = pRow.Value(pCatchPTable_S.Table.FindField("GBLAKES_ID"))
        dblCatchP(lonRowScroller, 1) = pRow.Value(pCatchPTable_S.Table.FindField("P")) 'units: 1 (microgram per litre) = 1.0 × 10-6 kilogram per (meter cubed)
        dblCatchP(lonRowScroller, 2) = pRow.Value(pCatchPTable_S.Table.FindField("LCOVDESC"))
        dblCatchP(lonRowScroller, 3) = pRow.Value(pCatchPTable_S.Table.FindField("Area")) 'units are square metres
        dblCatchP(lonRowScroller, 4) = dblCatchP(lonRowScroller, 1) / (dblCatchP(lonRowScroller, 3) / 10000) 'kg/ha
        dblCatchP(lonRowScroller, 5) = dblCatchP(lonRowScroller, 4) 'initialise with the read-in data
        Set pRow = pCursor.NextRow
        lonRowScroller = lonRowScroller + 1
    Wend
Else
    pQueryFilt.WhereClause = "Catch_Net = " & lonChosenNetwork
    lontblCatchPRecords = pCatchPTable.Table.RowCount(pQueryFilt)
    ReDim dblCatchP(lontblCatchPRecords - 1, 6)
    Set pCursor = pCatchPTable.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    lonRowScroller = 0
    While Not pRow Is Nothing
    'read P from each of the different landcover types for each catchment in the network
        dblCatchP(lonRowScroller, 0) = pRow.Value(intCatchP_GBLAKES_IDField)
        dblCatchP(lonRowScroller, 1) = pRow.Value(intCatchP_PField) 'units: 1 (microgram per litre) = 1.0 × 10-6 kilogram per (meter cubed)
        dblCatchP(lonRowScroller, 2) = pRow.Value(intCatchP_LCOVDESCField)
        dblCatchP(lonRowScroller, 3) = pRow.Value(intCatchP_AreaField) 'units square metres
        dblCatchP(lonRowScroller, 4) = dblCatchP(lonRowScroller, 1) / (dblCatchP(lonRowScroller, 3) / 10000) 'kg/ha
        dblCatchP(lonRowScroller, 5) = dblCatchP(lonRowScroller, 4) 'initialise with the read-in data
        Set pRow = pCursor.NextRow
        lonRowScroller = lonRowScroller + 1
    Wend
End If

'step through CatchNetRship() and sum the P from tblCatchP for each GBLAKES_ID in the network - note dblCatchP was read and populated above
For i = 0 To UBound(CatchNetRship, 1)
    For j = 0 To UBound(dblCatchP, 1)
        If CatchNetRship(i, 0) = dblCatchP(j, 0) Then
            CatchNetRship(i, 9) = CatchNetRship(i, 9) + dblCatchP(j, 1)
        End If
    Next
Next

'update CatchNetRship(lonChosenGBLAKES_ID, 9) to be the sum of arrayCatchPforChosenGBLAKES_ID(i,2) to take account of user entered values
For i = 0 To UBound(CatchNetRship, 1)
    If CatchNetRship(i, 0) = lonChosenGBLAKES_ID Then
        CatchNetRship(i, 9) = 0
        For j = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
            CatchNetRship(i, 9) = CatchNetRship(i, 9) + arrayCatchPforChosenGBLAKES_ID(j, 2)
        Next
        Exit For
    End If
Next

'enter the column headers for the listview output
lvwCatchmentInfo.ColumnHeaders.Clear
Dim intArrayColumnWidths3(4) As Integer
intArrayColumnWidths3(1) = 110
intArrayColumnWidths3(2) = 55
intArrayColumnWidths3(3) = 150
intArrayColumnWidths3(4) = 60
Dim strArrayColumnHeadings3(4) As String
strArrayColumnHeadings3(1) = "Site"
strArrayColumnHeadings3(2) = "P (kg)"
strArrayColumnHeadings3(3) = "Land cover"
strArrayColumnHeadings3(4) = "Area (m" & Chr(178) & ")"

For i = 1 To 4
   Set CatchmentInfoColumnHeaders1 = lvwCatchmentInfo.ColumnHeaders.Add()
   CatchmentInfoColumnHeaders1.Text = strArrayColumnHeadings3(i)
   CatchmentInfoColumnHeaders1.Width = intArrayColumnWidths3(i)
Next

lvwCatchmentInfo.ListItems.Clear
Dim dblAreaToUse As Double
Dim dblPToUse As Double
Dim blnChangeEffected As Boolean
blnChangeEffected = False
If Not cmdResolveAreaDifference.Enabled Then
    blnResolveDifferences = False
End If

'#######################################################################################
'Implement user entered values
'#######################################################################################
'save the CatchNetRship(intIndexSelectedGBLAKES_ID, 9) value for use below
Dim dblSelectedPSave As Double
dblSelectedPSave = CatchNetRship(intIndexSelectedGBLAKES_ID, 9)
'reset CatchNetRship(intIndexSelectedGBLAKES_ID,9) because we want to create it here using modified inputs as required
CatchNetRship(intIndexSelectedGBLAKES_ID, 9) = 0  '(i,9) is recalculated below
'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
'use the check boxes to control whether user entered values or the master values are used. Only modify items 5, 6 and 7, never the others
For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    If cmdModCatchmentInputs.Enabled = True Then
        If arrayCatchPforChosenGBLAKES_ID(i, 1) = strSelectedLandCoverType Then
            If chkChangeArea And optUserModified Then
                If IsNumeric(txtEnterNewArea) Then
                    If blnResolveDifferences Then
                        If arrayCatchPforChosenGBLAKES_ID(i, 1) = strSelectedLandCoverType Then 'only looking at the clicked land cover
                            blnChangeEffected = True    'don't want this to happen twice - below
                            'note that arrayCatchPforChosenGBLAKES_ID(i, 6) is assigned in cmdGetCatchmentInfo and only updated as needed
                            arrayCatchPforChosenGBLAKES_ID(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 6) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000)
                        Else
                            arrayCatchPforChosenGBLAKES_ID(i, 5) = CDbl(txtEnterNewArea)
                            arrayCatchPforChosenGBLAKES_ID(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 6) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000)
                        End If
                    Else
                        arrayCatchPforChosenGBLAKES_ID(i, 5) = CDbl(txtEnterNewArea)
                        arrayCatchPforChosenGBLAKES_ID(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 6) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000)
                    End If
                Else
                    MsgBox "Warning your entry for area is not a number", vbCritical
                    Exit Sub
                End If
            End If
            If (chkChangeP Or chkChangePforNetwork) And optUserModified Then
                If IsNumeric(txtEnterNewP) Then
                'the user entered P value for the selected land cover is implemented here, but only for the chosen GBLAKES_ID
                    If blnResolveDifferences And Not blnChangeEffected Then
                        If chkChangeArea And IsNumeric(txtEnterNewArea) Then
                                arrayCatchPforChosenGBLAKES_ID(i, 7) = CDbl(txtEnterNewP) * ((CDbl(txtEnterNewArea) + dblUserModifiedLCoverArea_difference) / 10000) 'if a user area is entered use it
                                arrayCatchPforChosenGBLAKES_ID(i, 6) = CDbl(txtEnterNewP)
                            Else
                                arrayCatchPforChosenGBLAKES_ID(i, 7) = CDbl(txtEnterNewP) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000) 'otherwise it's the read in
                                arrayCatchPforChosenGBLAKES_ID(i, 6) = CDbl(txtEnterNewP)
                        End If
                    Else
                        If chkChangeArea And IsNumeric(txtEnterNewArea) Then
                                arrayCatchPforChosenGBLAKES_ID(i, 7) = CDbl(txtEnterNewP) * (CDbl(txtEnterNewArea) / 10000) 'if a user area is entered use it
                                arrayCatchPforChosenGBLAKES_ID(i, 6) = CDbl(txtEnterNewP)
                            Else
                                arrayCatchPforChosenGBLAKES_ID(i, 7) = CDbl(txtEnterNewP) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000) 'otherwise it's the read in
                                arrayCatchPforChosenGBLAKES_ID(i, 6) = CDbl(txtEnterNewP)
                        End If
                    End If
                Else
                    MsgBox "Warning your entry for P is not a number", vbCritical
                    Exit Sub
                End If
            End If
        End If
    End If
    CatchNetRship(intIndexSelectedGBLAKES_ID, 9) = CatchNetRship(intIndexSelectedGBLAKES_ID, 9) + arrayCatchPforChosenGBLAKES_ID(i, 7)
Next

If cmdModCatchmentInputs.Enabled = False Then
    CatchNetRship(intIndexSelectedGBLAKES_ID, 9) = dblSelectedPSave
End If

Dim intNumMatchingInputs As Integer
intNumMatchingInputs = UBound(arrayCatchPforChosenGBLAKES_ID, 1)

dblSumLocalInputs = 0
For lonRowScroller = 0 To intNumMatchingInputs
    Set List_Item3 = lvwCatchmentInfo.ListItems.Add
    List_Item3 = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0)
    If optUserModified Then
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5)
    Else
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)
    End If
    If dblPToUse < 1 Then
            List_Item3.SubItems(1) = "0" & Format(dblPToUse, "#.0")
            dblSumLocalInputs = dblSumLocalInputs + dblPToUse
    Else
            List_Item3.SubItems(1) = Format(dblPToUse, "#.0")
            dblSumLocalInputs = dblSumLocalInputs + dblPToUse
    End If
    
    If optUserModified Then
        If (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)) _
        Or (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)) Then
            List_Item3.SubItems(2) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & " (mod.)" 'Text land cover for modified feature
        Else
            List_Item3.SubItems(2) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) 'Text land cover
        End If
    Else
        List_Item3.SubItems(2) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1)  'Text land cover
    End If
    If dblAreaToUse < 1 Then
            List_Item3.SubItems(3) = "0" & Format(dblAreaToUse, "#,#")
    Else
            List_Item3.SubItems(3) = Format(dblAreaToUse, "#,#")
    End If
    If dblAreaToUse = 0 Then
        List_Item3.SubItems(3) = 0
    End If
Next

'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
'Use information in arrayCatchPforChosenGBLAKES_ID(i,j) to update CatchNetRship(i, 9) = Sum P in tblCatchP for all of the landcover types for each GBLAKES_ID
'to implement chkChangePforNetwork

'#######################################################################################
'Implement user modification of the CatchP for a specfic land cover for the whole network
'#######################################################################################
Dim blnFoundInModArray As Boolean
Dim dblPtoAdd As Double
If chkChangePforNetwork.Value = True And optUserModified Then
    'step through CatchNetRship() and sum the P from tblCatchP for each GBLAKES_ID in the network - note dblCatchP was read and populated above
    For i = 0 To UBound(CatchNetRship, 1)
    CatchNetRship(i, 9) = 0 'reset Sum P in tblCatchP for all of the landcover types for each GBLAKES_ID
        For j = 0 To UBound(dblCatchP, 1)
            If CatchNetRship(i, 0) = dblCatchP(j, 0) Then
            blnFoundInModArray = False
            dblPtoAdd = 0
            'step through varModifiedLCoverCoeff() and if there is a match for varModifiedLCoverCoeff(i,0) then use varModifiedLCoverCoeff(i,1) for coeff.
            'note that the user entered coefficents are used on-the-fly - the dblCatchP table is not modified with the user entered coefficients
            For k = 0 To UBound(varModifiedLCoverCoeff, 1)
                'match the land covers to the store of modified land covers (for incremental changes in a single scenario)
                If varModifiedLCoverCoeff(k, 0) = dblCatchP(j, 2) Then
                    dblPtoAdd = varModifiedLCoverCoeff(k, 1)
                    blnFoundInModArray = True
                    Exit For
                End If
            Next
            If blnFoundInModArray Then
                CatchNetRship(i, 9) = CatchNetRship(i, 9) + (dblPtoAdd * dblCatchP(j, 3) / 10000)
            Else
                CatchNetRship(i, 9) = CatchNetRship(i, 9) + dblCatchP(j, 1)
            End If
            End If
        Next
    Next
End If

'for order 0, J = tblCatchP.P + s + x + y + 0 (note that R(up) and TP(up) are both 0 for order zero
'for each element the final term is (CatchNetRship(i, 3) - CatchNetRship(i, 2)) * TP(Up)

'The sigma R(Up) term is the sum of R upstream - so the upstream R = jAL + sigma R(Up)
'so for order 0, CatchNetRship(i, 3) = sum of the current and upstream local runoffs
'Each catchment only receives from its immediate upstream catchments - the cascading is determined that way. If A flows into B
'then B receives its TP * runoff / 1000000

'#######################################################################################
'Calculate by order - do all the 0s, then the 1s etc.
'Loop through all catchments - filtering by order 0 then 1 etc.
'for each catchment encountered look through the flows-into field of all other catchments for a match - if a match is found then increment J
'#######################################################################################
dblSumUpstream = 0  'this has been added on 07.11.2013 - it could have gone in below... it is reused in the export to CSV
If intMatchingGBLAKES_IDs = 0 Then
'calculate the simple case of a single unconnected catchment
    cmdCalcTP.Caption = "Calculate total P concentration (for " & lonChosenGBLAKES_ID & ")"
    cmdCalcTP.ControlTipText = "Calculate total P concentration or load (for " & lonChosenGBLAKES_ID & ")"
    CatchNetRship(0, 15) = 0
    CatchNetRship(0, 10) = CatchNetRship(0, 9) + CatchNetRship(0, 15) + CatchNetRship(0, 16) + CatchNetRship(0, 17) + CatchNetRship(0, 22) 'adding urban, rural and other point loads
    dblJSelectedCatchment = CatchNetRship(0, 10)
    'calculate TP - 'CatchNetRship(i, 3) = sum of the current and upstream local runoffs
    CatchNetRship(0, 14) = CatchNetRship(0, 10) * 1000000 / CatchNetRship(0, 3) 'sum P (mg) / Sum of current and upstream local runoffs?
    CatchNetRship(0, 14) = CatchNetRship(0, 14) / (1 + Sqr(CatchNetRship(0, 6)))
    CatchNetRship(0, 14) = CatchNetRship(0, 14) ^ CatchNetRship(0, 8)
    CatchNetRship(0, 14) = CatchNetRship(0, 14) * CatchNetRship(0, 7)
'#######################################################################################
'TP for unconnected catchments is not meaningful (i.e. intMatchingGBLAKES_IDs = 0), and will not be calculated or displayed
'#######################################################################################
    DoEvents
Else    'all catchments that connect to any other
    Dim intOrder As Integer
    intOrder = 0
    While intOrder <= maxOrder
    For i = 0 To UBound(lonOrderMatchArray())
        If CatchNetRship(i, 13) = intOrder Then
            If intOrder = 0 Then
                CatchNetRship(i, 15) = 0
                'dblSumUpstream = 0
                CatchNetRship(i, 10) = CatchNetRship(i, 9) + CatchNetRship(i, 16) + CatchNetRship(i, 17) + CatchNetRship(i, 22) 'adding urban, rural and point loads
                CatchNetRship(i, 14) = CatchNetRship(i, 10) * 1000000 / CatchNetRship(i, 3)
                CatchNetRship(i, 14) = CatchNetRship(i, 14) / (1 + Sqr(CatchNetRship(i, 6)))
                CatchNetRship(i, 14) = CatchNetRship(i, 14) ^ CatchNetRship(i, 8)
                CatchNetRship(i, 14) = CatchNetRship(i, 14) * CatchNetRship(i, 7)
                If CatchNetRship(i, 0) = lonChosenGBLAKES_ID Then
                    dblJSelectedCatchment = CatchNetRship(i, 10)
                End If
            Else    'the J is a combination of the current and those upstream that flow directly into it - note that doesn't mean intOrder - 1
                'calculate J(up) by finding the upstream matches
                For j = 0 To UBound(CatchNetRship, 1)
                    If CatchNetRship(j, 12) = CatchNetRship(i, 0) Then  'CatchNetRship(j, 0) flows into CatchNetRship(i, 0)
                        CatchNetRship(i, 15) = CatchNetRship(i, 15) + (CatchNetRship(j, 3) * CatchNetRship(j, 14) / 1000000)
                    End If
                Next
                CatchNetRship(i, 10) = CatchNetRship(i, 9) + CatchNetRship(i, 15) + CatchNetRship(i, 16) + CatchNetRship(i, 17) + CatchNetRship(i, 22) 'adding urban, rural and point loads 'adding urban and rural loads
                If CatchNetRship(i, 0) = lonChosenGBLAKES_ID Then
                    dblJSelectedCatchment = CatchNetRship(i, 10)
                    dblSumUpstream = CatchNetRship(i, 15)
                End If
            'now calculate the TP that will be used downstream
            'trap the entry of too large negative values - allowing some to model reduction in input, but if TP < 0 this would crash (but is trapped below)
                CatchNetRship(i, 14) = CatchNetRship(i, 10) * 1000000 / CatchNetRship(i, 3) 'input P and water residence time
                CatchNetRship(i, 14) = CatchNetRship(i, 14) / (1 + Sqr(CatchNetRship(i, 6)))
                If CatchNetRship(i, 14) >= 0 Then
                    CatchNetRship(i, 14) = CatchNetRship(i, 14) ^ CatchNetRship(i, 8)
                    CatchNetRship(i, 14) = CatchNetRship(i, 14) * CatchNetRship(i, 7)
                Else
                    MsgBox "The load you have entered is too low, resulting in a negative TP. Please enter a larger value.", vbCritical
                    Exit Sub
                End If
            End If
        End If
    Next
    intOrder = intOrder + 1
    Wend
End If

Set List_Item3 = lvwCatchmentInfo.ListItems.Add
Set List_Item3 = lvwCatchmentInfo.ListItems.Add

lonRowScroller = lonRowScroller + 1

'check the area against the original for user entered changes
'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
If optUserModified Then
    dblUserModifiedLCoverArea_difference = 0
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
        dblUserModifiedLCoverArea_difference = dblUserModifiedLCoverArea_difference - arrayCatchPforChosenGBLAKES_ID(i, 3) + arrayCatchPforChosenGBLAKES_ID(i, 5)
    Next
    If Abs(dblUserModifiedLCoverArea_difference) > 1 Then
        cmdResolveAreaDifference.Enabled = True
        cmdResolveAreaDifference.Visible = True
        cboResolveAreaDifference.Visible = True
        lblResolveAreaDifference.Visible = True
        If dblUserModifiedLCoverArea_difference < 0 Then
            lblAreaDifference.Caption = "Your total area is too low by " & Format(Abs(dblUserModifiedLCoverArea_difference), "#,#") & " m" & Chr(178)
        Else
            lblAreaDifference.Caption = "Your total area is too high by " & Format(dblUserModifiedLCoverArea_difference, "#,#") & " m" & Chr(178)
        End If
    Else
        lblAreaDifference.Caption = ""
        cmdResolveAreaDifference.Enabled = False
        cmdResolveAreaDifference.Visible = False
        cboResolveAreaDifference.Visible = False
        lblResolveAreaDifference.Visible = False
    End If
Else
    lblAreaDifference.Caption = ""
    cmdResolveAreaDifference.Enabled = False
    cmdResolveAreaDifference.Visible = False
    cboResolveAreaDifference.Visible = False
    lblResolveAreaDifference.Visible = False
End If
'#######################################################################################
'Begin creating the summary outputs to the screen/form
'#######################################################################################
If dblJSelectedCatchment < 1 Then
    List_Item3 = "Sum P (kg) = 0" & Format(dblJSelectedCatchment, "#.0")
Else
    List_Item3 = "Sum P (kg) = " & Format(dblJSelectedCatchment, "#.0")
End If
If dblSumLocalInputs < 1 Then
    List_Item3.SubItems(2) = "Sum land cover inputs (kg P) = 0" & Format(dblSumLocalInputs, "#.0")
Else
    List_Item3.SubItems(2) = "Sum land cover inputs (kg P) = " & Format(dblSumLocalInputs, "#.0")
End If
Set List_Item3 = lvwCatchmentInfo.ListItems.Add
List_Item3 = ""
If dblSumUpstream < 1 Then
    List_Item3.SubItems(2) = "Sum upstream inputs (kg P) = 0" & Format(dblSumUpstream, "#.0")
Else
    List_Item3.SubItems(2) = "Sum upstream inputs (kg P) = " & Format(dblSumUpstream, "#.0")
End If
'#######################################################################################
'Output the rural and urban sewage for the selected GBLAKES_ID
'#######################################################################################
Set List_Item3 = lvwCatchmentInfo.ListItems.Add
Set List_Item3 = lvwCatchmentInfo.ListItems.Add 'add a couple of empty lines for spacing
List_Item3 = "Sewage Loads"
List_Item3.SubItems(2) = "(for site " & lonChosenGBLAKES_ID & ")"

For i = 0 To UBound(varCatchmentSewage, 1)
    If CatchNetRship(intIndexSelectedGBLAKES_ID, 0) = varCatchmentSewage(i, 0) Then
        Set List_Item3 = lvwCatchmentInfo.ListItems.Add
        If CatchNetRship(intIndexSelectedGBLAKES_ID, 16) < 1 Then
            List_Item3 = "Urban (kg P) = 0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0")
        Else
            List_Item3 = "Urban (kg P) = " & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0")
        End If
        If CatchNetRship(intIndexSelectedGBLAKES_ID, 17) < 1 Then
            List_Item3.SubItems(2) = "Rural (kg P) = 0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0")
        Else
            List_Item3.SubItems(2) = "Rural (kg P) = " & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0")
        End If
        Set List_Item3 = lvwCatchmentInfo.ListItems.Add
        If CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = 0 Then
            List_Item3 = "Urban population = 0"
        Else
            List_Item3 = "Urban population = " & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 18), "#")
        End If
        If CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = 0 Then
            List_Item3.SubItems(2) = "Rural population = 0"
        Else
            List_Item3.SubItems(2) = "Rural population = " & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 19), "#")
        End If
    End If
Next

'#######################################################################################
'Create the point source output
'#######################################################################################
For i = 0 To UBound(varPointSource, 1)
    If varPointSource(i, 0) = lonChosenGBLAKES_ID Then
        If varPointSource(i, 2) <> 0 Then
            If i = 0 Then 'header lines...
                Set List_Item3 = lvwCatchmentInfo.ListItems.Add
                Set List_Item3 = lvwCatchmentInfo.ListItems.Add 'add a couple of empty lines for spacing
                List_Item3 = "Point sources"
                List_Item3.SubItems(2) = "(for site " & lonChosenGBLAKES_ID & ")"
            End If
            Set List_Item3 = lvwCatchmentInfo.ListItems.Add
            List_Item3 = varPointSource(i, 1) & " (kg P)"
            List_Item3.SubItems(1) = varPointSource(i, 2)
        End If
    End If
Next

lvwCatchmentRelationships1.Height = 0
lvwCatchmentRelationships1.Visible = False
lvwCatchmentRelationships2.Visible = True
lstViewSupplement.Visible = True
lvwCatchmentRelationships2.Height = intTopListviewBoxesHeight
lvwCatchmentRelationships2.Top = intTopListviewBoxesTop
'output to multicolumn listview use the index as the selection ID
lvwCatchmentRelationships2.ColumnHeaders.Clear
Dim intArrayColumnWidths(16) As Integer
intArrayColumnWidths(1) = 32
intArrayColumnWidths(2) = 78
intArrayColumnWidths(3) = 35
intArrayColumnWidths(4) = 108
intArrayColumnWidths(5) = 30
intArrayColumnWidths(6) = 42
intArrayColumnWidths(7) = 45
intArrayColumnWidths(8) = 45
intArrayColumnWidths(9) = 72
intArrayColumnWidths(10) = 70
intArrayColumnWidths(11) = 74
intArrayColumnWidths(12) = 75
intArrayColumnWidths(13) = 86
intArrayColumnWidths(14) = 82
intArrayColumnWidths(15) = 86
intArrayColumnWidths(16) = 84
Dim strArrayColumnHeadings(16) As String
strArrayColumnHeadings(1) = "Status"
strArrayColumnHeadings(2) = "SEPA RAG & status"
strArrayColumnHeadings(3) = "Site"
strArrayColumnHeadings(4) = "PLUS+ modelled RAG & Site"
strArrayColumnHeadings(5) = "Order"
strArrayColumnHeadings(6) = "TP (" & Chr(181) & "g/l)"
strArrayColumnHeadings(7) = "J (kg)"
strArrayColumnHeadings(8) = "Ref. type"
strArrayColumnHeadings(9) = "Cap. to down, TP"
strArrayColumnHeadings(10) = "Cap. to down, J" '(kg)"
strArrayColumnHeadings(11) = "Cap. to upgrd, TP"
strArrayColumnHeadings(12) = "Cap. to upgrd, J" '(kg)"
strArrayColumnHeadings(13) = "Meas.Cap. down, TP"   'SEPA point measurement data comparison
strArrayColumnHeadings(14) = "Meas.Cap. down, J" '(kg)"
strArrayColumnHeadings(15) = "Meas.Cap. upgrd, TP"
strArrayColumnHeadings(16) = "Meas.Cap. upgrd, J" '(kg)"
For i = 1 To 16
   Set CatchmentRelateColumnHeaders2 = lvwCatchmentRelationships2.ColumnHeaders.Add()
   CatchmentRelateColumnHeaders2.Text = strArrayColumnHeadings(i)
   CatchmentRelateColumnHeaders2.Width = intArrayColumnWidths(i)
Next

'Add the SEPA annual assessment from pSEPAmonitoringArray (note that this is not the measured status, that follows)
'pSEPAmonitoringArray(X, 0) = WATER_BODY_ID ... etc. WATER_BODY_NAME,CLASSIFICATION_YEAR,STATUS"
'Find the waterbody within the array so use the GBLakes_WBID_lookup to translate from the
'GBLAKES_ID (GB_WB_ID) to the WFD_WB_ID (as used in SEPA_Loch_WB_classification)
'use a function to return the SEPA P classification which in turn calls another function which translates between the IDs

Dim strToolTipText As String
Dim strColourReturned As String
lvwCatchmentRelationships2.ListItems.Clear
Dim dblArrayUpgradeDowngrade(4) As Double
Dim dblArrayMeasUpgradeDowngrade(4) As Double
Dim dblTestUpDown As Double
Dim dblUpgradeMark As Double
Dim dblDowngradeMark As Double
dblDowngradeMark = 0
dblUpgradeMark = 0
Dim dblTemp As Double
'add the information derived from the SEPA monitoring from the table selected in cboClassConcStat, that is pSEPAClassConcStatTable
'this is another capacity calculation, however, instead of the modelled concentration (i.e. the output from CalcTP) the value in
'pSEPAClassConcStatArray(lonCounter, 2) is used

'    pSEPAClassConcStatArray(lonCounter, 0) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("WATER_BODY_ID"))
'    pSEPAClassConcStatArray(lonCounter, 1) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("YEAR_"))
'    pSEPAClassConcStatArray(lonCounter, 2) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("POINT_CLASSIFICATION_RESULT"))
'    pSEPAClassConcStatArray(lonCounter, 3) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("CLASS_ID"))
'note that CLASS_ID is 1 - 5, corresponding to High -> poor

'The measured values do not over-write the modelled values throughout the calculation (the measured concentrations are not used in modelling)
'We are just calculating the amount and conc. to upgrade/downgrade for a given measured value and sending the output to the window

'format of dblArrayMeasUpgradeDowngrade is (1) = Cap to TP downgrade, (2) = Cap to J down, (3) = Cap to TP up, (4) = Cap to J up
'however, (2) and (4) are not actually populated, they are calculated in dblTestUpDown = ...
'This is not a scenario, so these measured values are not flushed through the network - it is just an alternate view of a catchment
'whether in scenario mode or not - but only for those lochs with a WBID
dblSEPA_meas_conc = 0
Dim strToolTipText2 As String
Dim dblReturnedWFD_WB_ID As Double
'strDiscrepancyInClasses = ""' if we reset this then it messes up batchmode
For j = 0 To UBound(CatchNetRship, 1)

'modification on 7th September 2016 to back calculate the loads associated with RAG points ->
'need to run on all catchments
Dim pQueryFilt3 As IQueryFilter2
Set pQueryFilt3 = New QueryFilter
pQueryFilt.SubFields = "GBLAKES_ID,Reference_Type,HighGood_P,GoodModerate_P,ModeratePoor_P,PoorBad_P"

Dim lon_GBLAKES_ID As Long
lon_GBLAKES_ID = CatchNetRship(j, 0)
Dim pCursor3 As ICursor
Dim pRow3 As IRow
pQueryFilt3.WhereClause = "SiteCode = " & lon_GBLAKES_ID
Set pCursor3 = pTPBreakPoints.Table.Search(pQueryFilt3, False)
Set pRow3 = pCursor3.NextRow
Dim intRowCount As Integer
Dim dblRA_TP As Double
Dim dblRA_J As Double
Dim dblAG_TP As Double
Dim dblAG_J As Double
intRowCount = 0
While Not pRow3 Is Nothing
'note that debug only outputs the first 200 lines...
'do high
    dblRA_TP = pRow3.Value(intHighGood_PField) - ((pRow3.Value(intHighGood_PField) - 0) * 0.03)
    dblAG_TP = pRow3.Value(intHighGood_PField) - ((pRow3.Value(intHighGood_PField) - 0) * 0.2)
    dblRA_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblRA_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    dblAG_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblAG_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    Debug.Print lon_GBLAKES_ID & ":High:Red/Amber TP:" & dblRA_TP & ":load:" & dblRA_J & ":High:Amber/GreenTP:" & dblAG_TP & ":load:" & dblAG_J
'do good
    dblRA_TP = pRow3.Value(intGoodModerate_PField) - ((pRow3.Value(intGoodModerate_PField) - pRow3.Value(intHighGood_PField)) * 0.03)
    dblAG_TP = pRow3.Value(intGoodModerate_PField) - ((pRow3.Value(intGoodModerate_PField) - pRow3.Value(intHighGood_PField)) * 0.2)
    dblRA_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblRA_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    dblAG_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblAG_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    Debug.Print lon_GBLAKES_ID & ":Good:Red/Amber TP:" & dblRA_TP & ":load:" & dblRA_J & ":Good:Amber/Green TP:" & dblAG_TP & ":load:" & dblAG_J
'do moderate
    dblRA_TP = pRow3.Value(intModeratePoor_PField) - ((pRow3.Value(intModeratePoor_PField) - pRow3.Value(intGoodModerate_PField)) * 0.03)
    dblAG_TP = pRow3.Value(intModeratePoor_PField) - ((pRow3.Value(intModeratePoor_PField) - pRow3.Value(intGoodModerate_PField)) * 0.2)
    dblRA_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblRA_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    dblAG_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblAG_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
'    Debug.Print lon_GBLAKES_ID & ", Moderate: Red/Amber TP: " & dblRA_TP & " and load: " & dblRA_J & ", Moderate: Amber/Green TP: " & dblAG_TP & " and load: " & dblAG_J
'do poor
    dblRA_TP = pRow3.Value(intPoorBad_PField) - ((pRow3.Value(intPoorBad_PField) - pRow3.Value(intModeratePoor_PField)) * 0.03)
    dblAG_TP = pRow3.Value(intPoorBad_PField) - ((pRow3.Value(intPoorBad_PField) - pRow3.Value(intModeratePoor_PField)) * 0.2)
    dblRA_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblRA_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
    dblAG_J = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblAG_TP / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
'    Debug.Print lon_GBLAKES_ID & ", Poor: Red/Amber TP: " & dblRA_TP & " and load: " & dblRA_J & ", Poor: Amber/Green TP: " & dblAG_TP & " and load: " & dblAG_J
    Set pRow3 = pCursor3.NextRow
Wend
'<- end of modification on 7th September 2016 to back calculate the loads associated with RAG points

    dblReturnedWFD_WB_ID = ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0)))
    For i = 0 To 4
        dblArrayUpgradeDowngrade(i) = 0
        dblArrayMeasUpgradeDowngrade(i) = 0
    Next
    Set List_Item2 = lvwCatchmentRelationships2.ListItems.Add
    If CatchNetRship(j, 0) = lonChosenGBLAKES_ID Then
        List_Item2.Bold = True
        List_Item2.Text = "*"
    Else
        List_Item2.Text = ""
    End If
    'send the GBLAKES_ID and a concentration to ColourToDisplay to calculate the colour to display.
    'format of dblArrayUpgradeDowngrade is (1) = Cap to TP downgrade, (2) = Cap to J down, (3) = Cap to TP up, (4) = Cap to J up
    'J to down/up is a function of TP down/up, volume,
    
    strColourReturned = ColourToDisplay(CLng(CatchNetRship(j, 0)), CDbl(CatchNetRship(j, 14)), strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark, True)
    List_Item2.SmallIcon = strColourReturned
    List_Item2.ToolTipText = strToolTipText
    CatchNetRship(j, 21) = strToolTipText

    Dim varSplitTemp As Variant
    Dim strTPBreakPointStatusMeasValue As String
    If dblReturnedWFD_WB_ID <> 0 Then
    'insert RAG for SEPA measured status - compare the measured value with the TPBreakPoints value
    'if the derived water quality does not correspond to the observed water quality then output to a window.
        List_Item2.SubItems(1) = dblReturnedWFD_WB_ID & " : " & ReturnSEPA_Status(CatchNetRship(j, 0)) 'this is the SEPA supplied status, no calculation
        For k = 0 To UBound(pSEPAClassConcStatArray, 1)
            If pSEPAClassConcStatArray(k, 0) = dblReturnedWFD_WB_ID Then
'DisplayRAG returns the RAG colour in strToolTipText as a colour number
                DisplayRAG CLng(CatchNetRship(j, 0)), CDbl(pSEPAClassConcStatArray(k, 2)), strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark
                If strToolTipText Like "*selected site" Then
                    varSplitTemp = Split(strToolTipText, " ")
                    strToolTipText = varSplitTemp(0)
                End If
                Select Case strToolTipText
                    Case "255"
                    strToolTipText = "RAG: Red"
                    List_Item2.ListSubItems.Item(1).ReportIcon = "Bad" 'Red
                    List_Item2.ListSubItems.Item(1).ToolTipText = "RAG: Red"
                    Case "33023"
                    strToolTipText = "RAG: Amber"
                    List_Item2.ListSubItems.Item(1).ReportIcon = "Poor" 'Amber
                    List_Item2.ListSubItems.Item(1).ToolTipText = "RAG: Amber"
                    Case "65280"
                    strToolTipText = "RAG: Green"
                    List_Item2.ListSubItems.Item(1).ReportIcon = "Good" 'Green
                    List_Item2.ListSubItems.Item(1).ToolTipText = "RAG: Green"
                End Select
'    pSEPAClassConcStatArray(lonCounter, 0) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("WATER_BODY_ID"))
'    pSEPAClassConcStatArray(lonCounter, 1) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("YEAR_"))
'    pSEPAClassConcStatArray(lonCounter, 2) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("POINT_CLASSIFICATION_RESULT"))
'    pSEPAClassConcStatArray(lonCounter, 3) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("CLASS_ID"))
                'calculate the TPBreakPoint status for the measured P - always use TPBreakPoint for status calculations
                strTPBreakPointStatusMeasValue = ColourToDisplay(CLng(CatchNetRship(j, 0)), CDbl(pSEPAClassConcStatArray(k, 2)), strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark, False)
                If strTPBreakPointStatusMeasValue = "Bad" Then '"Bad" is always green (email communication from J. Bowes
                    List_Item2.ListSubItems.Item(1).ReportIcon = "Good" 'Green
                    List_Item2.ListSubItems.Item(1).ToolTipText = "RAG: Green, Status: Bad"
                End If
                If strTPBreakPointStatusMeasValue <> ReturnHighFor1(pSEPAClassConcStatArray(k, 3)) Then 'output a message if the status in the SEPA table does not match the status derived from the measurement (strTPBreakPointStatusMeasValue)
                    strDiscrepancyInClasses = strDiscrepancyInClasses & "TPBreakPoints status of measured value of " _
                    & pSEPAClassConcStatArray(k, 0) & ": Status derived from conc.: " & strTPBreakPointStatusMeasValue & ", SEPA table status: " _
                    & ReturnHighFor1(pSEPAClassConcStatArray(k, 3)) & "(" & pSEPAClassConcStatArray(k, 0) _
                    & ", " & pSEPAClassConcStatArray(k, 2) & ", " & pSEPAClassConcStatArray(k, 3) & ")" & vbCrLf
                    'this message is output to the immediate window at the end of the For..Next loop, and is saved as a text file if a report is created
                    List_Item2.SubItems(1) = List_Item2.SubItems(1) & ", TPBreakPoints = " & Left(strTPBreakPointStatusMeasValue, 1)
                End If
            End If
        Next
        
    Else
        List_Item2.SubItems(1) = ""
    End If
    'SubItems(2) is moved down a little for efficiency of code
    List_Item2.SubItems(3) = ReturnSitename(CLng(CatchNetRship(j, 0)))
    
    strColourReturned = ColourToDisplay(CLng(CatchNetRship(j, 0)), CDbl(CatchNetRship(j, 14)), strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark, True)
    DisplayRAG CLng(CatchNetRship(j, 0)), CDbl(CatchNetRship(j, 14)), strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark
    If strToolTipText Like "*selected site" Then
        varSplitTemp = Split(strToolTipText, " ")
        strToolTipText = varSplitTemp(0)
    End If
    If strColourReturned = "Bad" Then '"Bad" is always green (email communication from J. Bowes)
        strToolTipText = "RAG: Green"
        List_Item2.ListSubItems.Item(3).ReportIcon = "Good" 'Green
    Else
        Select Case strToolTipText
        Case "255"
        strToolTipText = "RAG: Red"
        List_Item2.ListSubItems.Item(3).ReportIcon = "Bad" 'Red
        Case "33023"
        strToolTipText = "RAG: Amber"
        List_Item2.ListSubItems.Item(3).ReportIcon = "Poor" 'Amber
        Case "65280"
        strToolTipText = "RAG: Green"
        List_Item2.ListSubItems.Item(3).ReportIcon = "Good" 'Green
        End Select
    End If
    List_Item2.ListSubItems.Item(3).ToolTipText = strToolTipText
    
    List_Item2.SubItems(4) = "  " & lonOrderMatchArray(j)
    List_Item2.ListSubItems.Item(4).ToolTipText = CatchNetRship(j, 1)
    If CatchNetRship(j, 0) = lonChosenGBLAKES_ID Then
        List_Item2.ListSubItems.Item(4).ToolTipText = "Chosen body"
    End If
    Select Case List_Item2.ListSubItems.Item(4).ToolTipText
    Case "Chosen body"
    
    Case "Is Downstream from"
        List_Item2.ListSubItems.Item(4).ToolTipText = List_Item2.ListSubItems.Item(4).ToolTipText & " " & lonChosenGBLAKES_ID
    Case "Separate Branch"
        List_Item2.ListSubItems.Item(4).ToolTipText = List_Item2.ListSubItems.Item(4).ToolTipText & " from " & lonChosenGBLAKES_ID
    Case "Is Upstream of"
        List_Item2.ListSubItems.Item(4).ToolTipText = List_Item2.ListSubItems.Item(4).ToolTipText & " " & lonChosenGBLAKES_ID
    End Select
    'add the SEPA measured concentrations as a tooltip
'modifications done on 8th December 2015
'the tooltip to display the measured concentration wasn't appearing for standalone catchments
'so I removed the filter below and updated the if statement to add And dblReturnedWFD_WB_ID <> 0
    'If intMatchingGBLAKES_IDs = 0 Then
        If CatchNetRship(j, 14) < 1 Then
            List_Item2.SubItems(5) = "0" & Format(CatchNetRship(j, 14), "#.00")
        Else
            List_Item2.SubItems(5) = Format(CatchNetRship(j, 14), "#.00")
        End If
    'Else
        'get the tooltip text string
        strToolTipText2 = ""
        List_Item2.ListSubItems.Item(5).ToolTipText = strToolTipText2
        For k = 0 To UBound(pSEPAClassConcStatArray, 1)
            If pSEPAClassConcStatArray(k, 0) = dblReturnedWFD_WB_ID And dblReturnedWFD_WB_ID <> 0 Then
                If pSEPAClassConcStatArray(k, 2) < 1 Then
                    strToolTipText2 = "Meas: 0" & Format(pSEPAClassConcStatArray(k, 2), "#.0")
                    List_Item2.ListSubItems.Item(5).ToolTipText = strToolTipText2
                Else
                    strToolTipText2 = "Meas: " & Format(pSEPAClassConcStatArray(k, 2), "#.0")
                    List_Item2.ListSubItems.Item(5).ToolTipText = strToolTipText2
                End If
            End If
        Next
    'End If
    If CatchNetRship(j, 10) < 1 Then
        List_Item2.SubItems(6) = "0" & Format(CatchNetRship(j, 10), "#.0")
    Else
        List_Item2.SubItems(6) = Format(CatchNetRship(j, 10), "#.0")
    End If
    List_Item2.SubItems(7) = strTPBreakPointsRefType
    CatchNetRship(j, 20) = strTPBreakPointsRefType
    'if we are already 'bad' this is adjusted later
    List_Item2.SubItems(8) = Format(dblArrayUpgradeDowngrade(1), "0.0") 'cap to downgrade
    
    'Calculate capacity to downgrade J
    'TP is in ug/l, which equates to mg/m3. Runoff is in m3yr-1 so need to divide output by 1,000,000 to get kg
    'reverse calculate the concentration from dblDowngradeMark - the exponent term means that this is a non-linear relationship
    If dblDowngradeMark <> 9999 Then
            'J of break point = sum curr_and_upstr_runoff*(1 + sqr(Tw)))*((downgrd_brk_pt/OECD-a) ^ 1/OECD-b
            'calculate the J (kg) corresponding to the breakpoint
            dblTestUpDown = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblDowngradeMark / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
            'and remove the current TP to calculate the capacity
            dblTestUpDown = dblTestUpDown - CatchNetRship(j, 10)
    End If
    List_Item2.SubItems(9) = Format(dblTestUpDown, "0.0") 'Cap. to down, J
    If dblUpgradeMark <> 9999 Then
        List_Item2.SubItems(10) = Format(dblArrayUpgradeDowngrade(3) * -1, "0.0") 'multiplying  by -1 to indicate to the user that P must be subtracted to upgrade
        dblTestUpDown = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblUpgradeMark / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
        'and remove the current TP
        dblTestUpDown = dblTestUpDown - CatchNetRship(j, 10) 'CatchNetRship(j, 10) is J (kg), sum of the current and immediate upstream P (total P in the loch in kg)
    End If
    List_Item2.SubItems(11) = Format(dblTestUpDown, "0.0")
    If strColourReturned = "High" Then 'cannot upgrade beyond high
        List_Item2.SubItems(10) = "N/A is High"
        List_Item2.SubItems(11) = "N/A is High"
    ElseIf strColourReturned = "Bad" Then 'cannot downgrade beyond bad
        List_Item2.SubItems(8) = "N/A is Bad"
        List_Item2.SubItems(9) = "N/A is Bad"
    End If
    
    List_Item2.SubItems(2) = CatchNetRship(j, 0)
    If dblReturnedWFD_WB_ID <> 0 Then

'############ SEPA MEASURED CONC. and the resulting CAPACITY #################
'calculate the measured capacity to up/downgrade for these items
        For k = 0 To UBound(pSEPAClassConcStatArray, 1)

            If pSEPAClassConcStatArray(k, 0) = ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0))) Then
                dblSEPA_meas_conc = pSEPAClassConcStatArray(k, 2)
'compare it to the relevant breakpoints - are comparing to the TPBreakPoints data
'Have now produced a tool to allow the user to update TPBreakPoints table, so it is the responsibility of the user to ensure the tables are consistent.

'"ColourToDisplay" also sets up the Additional Results tab which isn't wanted here so the boolean at the end instructs it not to update this tab.

                ColourToDisplay CLng(CatchNetRship(j, 0)), dblSEPA_meas_conc, strToolTipText, dblArrayUpgradeDowngrade, dblDowngradeMark, dblUpgradeMark, False

'#####################################################################################################################################
'Note that this will also calculate the relevant breakpoints so that the output capacity will be based on the calculated break points
'and not the status read in from any table. i.e. the script will calculate which breakpoint to use - an alternative approach is
'required for calculating "SEPA monitored RAG" - however, I should compare the stated SEPA status (H, G etc.) to the calculated status
'and display a warning to a user if there are discrepancies.
'#####################################################################################################################################

'format of dblArrayUpgradeDowngrade is (1) = Cap to TP downgrade, (2) = Cap to J down, (3) = Cap to TP up, (4) = Cap to J up
                
'Note that the calculation for capacity incorporates T - sum of local and u/stream runoff.
'calculate J(kg) that corresponds to measured TP(ug/l)
                dblJ_for_meas_TP = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblSEPA_meas_conc / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
                CatchNetRship(j, 25) = dblSEPA_meas_conc
                CatchNetRship(j, 26) = dblJ_for_meas_TP


                List_Item2.ListSubItems.Item(6).ToolTipText = "Calc. J for SEPA conc.: " & Format(dblJ_for_meas_TP, "#.0")
                If dblDowngradeMark <> 9999 Then
                    List_Item2.SubItems(12) = Format(dblArrayUpgradeDowngrade(1), "0.0")
                'calculate J(kg) that corresponds to downgrade
                    dblTestUpDown = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblDowngradeMark / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
                    'and remove the J(kg) calculated for the measured TP
                    dblTestUpDown = dblTestUpDown - dblJ_for_meas_TP
                End If
                List_Item2.SubItems(13) = Format(dblTestUpDown, "0.0") 'Cap. to down, J
    
                If dblUpgradeMark <> 9999 Then
                    List_Item2.SubItems(14) = Format(dblArrayUpgradeDowngrade(3) * -1, "0.0") 'multiplying  by -1 to indicate to the user that P must be subtracted
                'calculate J(kg) that corresponds to upgrade
                    dblTestUpDown = (CatchNetRship(j, 3) * (1 + Sqr(CatchNetRship(j, 6)))) * ((dblUpgradeMark / CatchNetRship(j, 7)) ^ (1 / CatchNetRship(j, 8)) / 1000000)
                    'and remove the J(kg) calculated for the measured TP
                    dblTestUpDown = dblTestUpDown - dblJ_for_meas_TP
                End If
                List_Item2.SubItems(15) = Format(dblTestUpDown, "0.0")  'this can be the amount to upgrade from a SEPA measured medium to a PLUS+ high
    
                If dblUpgradeMark = 9999 Then 'cannot upgrade beyond high
                    List_Item2.SubItems(14) = "N/A is High"
                    List_Item2.SubItems(15) = "N/A is High"
                ElseIf dblDowngradeMark = 9999 Then 'cannot downgrade beyond bad
                    List_Item2.SubItems(12) = "N/A is Bad"
                    List_Item2.SubItems(13) = "N/A is Bad"
                End If
            End If
        Next
    End If
Next

'If strDiscrepancyInClasses <> "" Then
'    Debug.Print strDiscrepancyInClasses
'End If

lvwCatchmentInfo.Visible = True
lvwCatchmentInfo.Height = intTopListviewBoxesHeight + 12
lvwCatchmentInfo.Top = 147
lvwCatchmentInfo.HideSelection = True
If chkChangePforNetwork.Value = True Then
    cmdModCatchmentInputs.Caption = "Modify inputs" & vbCrLf & "(for whole network)"
Else
    cmdModCatchmentInputs.Caption = "Modify inputs" & vbCrLf & "(for " & lonChosenGBLAKES_ID & ")"
End If
Frame_Modify_Load.Visible = True
frameModifyInputs.Visible = True
DisplaySewageInfo
'initialise report output in the application window
If blnDataLoadedFromAScenario Then
    If ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) <> 0 Then
        txtMapTitle.Text = "PLUS+ report for scenario " & lonSelectedScenario & ", site " & lonChosenGBLAKES_ID & " (" & ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) & "), " & ReturnSitename(lonChosenGBLAKES_ID) & "."
        Else
        txtMapTitle.Text = "PLUS+ report for scenario " & lonSelectedScenario & ", site " & lonChosenGBLAKES_ID & ", " & ReturnSitename(lonChosenGBLAKES_ID) & "."
    End If
Else
    If ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) <> 0 Then
        txtMapTitle.Text = "PLUS+ report for site " & lonChosenGBLAKES_ID & " (" & ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) & "), " & ReturnSitename(lonChosenGBLAKES_ID) & "."
        Else
        txtMapTitle.Text = "PLUS+ report for site " & lonChosenGBLAKES_ID & ", " & ReturnSitename(lonChosenGBLAKES_ID) & "."
    End If
End If

If optReportOnScenario Then
    txtMapTitle.Text = Left(txtMapTitle.Text, Len(txtMapTitle.Text) - 1) & " - Scenario"
End If
chkProduceResultsTable.Enabled = True
chkProduceResultsCSV.Enabled = True

DoEvents

End Sub
Private Sub cmdCreateReport_Click()
'Create the output report as a JPEG or PDF map

'Dim pMxDoc As IMxDocument
Dim pActiveView As IActiveView
Dim pExport As IExport
Dim pPixelBoundsEnv As IEnvelope
Dim exportRECT As tagRECT
Dim iOutputResolution As Integer
Dim iScreenResolution As Integer
Dim hdc As Long
Dim i As Integer

Set pMxDoc = ThisDocument

'get the JPEG quality from the menu box
Dim pExportJPEG As IExportJPEG
Set pExportJPEG = New ExportJPEG
pExportJPEG.Quality = CInt(txtJPEGQuality.Text)

If optJPEG Then
    Set pExport = New ExportJPEG
    Set pExport = pExportJPEG
Else
    Set pExport = New ExportPDF
End If

'get the status of chkProduceResultsTable and chkProduceResultsCSV and re-instate them before exporting
Dim blnchkProduceResultsTable As Boolean
Dim blnchkProduceResultsCSV As Boolean
blnchkProduceResultsTable = chkProduceResultsTable.Value
blnchkProduceResultsCSV = chkProduceResultsCSV.Value

'check a valid name has been input and it doesn't end in "\"
If Right(txtOutputReport.Text, 1) = "\" Then
    MsgBox "Warning you don't appear to have entered a root file name for your output maps and report. Please do so." _
    & vbCrLf & "Please also make sure you have a root file name for the text output files.", vbCritical
    Exit Sub
End If
If chkProduceResultsCSV Then 'check a valid name has been input and it doesn't end in "\"
    If Right(txtOutputFile.Text, 1) = "\" Then
        MsgBox "Warning you don't appear to have entered a root file name for your text output files. Please do so.", vbCritical
        Exit Sub
    End If
End If

'ensure a JPEG has a jpg extension and a PDF a pdf extension
If optJPEG Then
    If Right(txtOutputReport.Text, 4) = ".pdf" Then
        txtOutputReport.Text = Left(txtOutputReport.Text, (Len(txtOutputReport.Text) - 4))
    End If
    If Right(txtOutputReport.Text, 4) <> ".jpg" Then
        txtOutputReport.Text = txtOutputReport.Text & ".jpg"
    End If
Else
    If Right(txtOutputReport.Text, 4) = ".jpg" Then
        txtOutputReport.Text = Left(txtOutputReport.Text, (Len(txtOutputReport.Text) - 4))
    End If
    If Right(txtOutputReport.Text, 4) <> ".pdf" Then
        txtOutputReport.Text = txtOutputReport.Text & ".pdf"
    End If
End If

If chkProduceResultsCSV Then
    'ensure the text output file is a text extension
    If Left(Right(txtOutputFile.Text, 4), 1) = "." Then
            txtOutputFile.Text = Left(txtOutputFile.Text, (Len(txtOutputFile.Text) - 4)) & ".txt"
        Else
            txtOutputFile.Text = txtOutputFile.Text & ".txt"
    End If
End If

chkProduceResultsTable.Value = blnchkProduceResultsTable
chkProduceResultsCSV.Value = blnchkProduceResultsCSV

'create the six text output file names
If chkProduceResultsCSV Then
    Dim strArrayTextFileNames(5) As String
    Dim strTempName As String
    strTempName = Left(txtOutputFile.Text, (Len(txtOutputFile.Text) - 4))
    strArrayTextFileNames(0) = strTempName & "_CatchRelationships.txt"
    strArrayTextFileNames(1) = strTempName & "_CatchP.txt"
    strArrayTextFileNames(2) = strTempName & "_LandCoverP.txt"
    strArrayTextFileNames(3) = strTempName & "_Summary.txt"
    strArrayTextFileNames(4) = strTempName & "_PointSource.txt"
    strArrayTextFileNames(5) = strTempName & "_Capacity.txt"
    
    'check the names are suitable
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim j As Integer
    For j = 0 To UBound(strArrayTextFileNames)
        If Not fso.FileExists(strArrayTextFileNames(j)) Then
            If fso.FolderExists(strArrayTextFileNames(j)) Then
                MsgBox strArrayTextFileNames(j) & " is an existing folder/path that can't be overwritten." & vbCrLf & _
                        "Please enter your output file", vbInformation, "Warning"
                Exit Sub
            End If
        End If
    Next
End If

chkProduceResultsTable.Value = blnchkProduceResultsTable
chkProduceResultsCSV.Value = blnchkProduceResultsCSV

'Switch to layout view
If Not pMxDoc.ActiveView Is pMxDoc.PageLayout Then
    Set pMxDoc.ActiveView = pMxDoc.PageLayout
End If
Set pActiveView = pMxDoc.ActiveView

'implement landscape/portrait option
'cannot assume that the map is in the correct orientation for either option...
'get the map scale and zoom out to a round number (if required)
Dim pPageLayout As IPageLayout3
Dim pMapFrame As IMapFrame
Dim CurrentMapScale As Long
Dim pPrinter As IPrinter
Dim pClone As IClone

Set pPageLayout = pMxDoc.PageLayout
If optMapOfNetwork Then
    optNetwork = True
    cmdZoomToSelected_Click
End If

If optMapOfCatchment Then
    optCatchment = True
    cmdZoomToSelected_Click
End If
Set pActiveView = pPageLayout

Set pGraphicsContainer = pPageLayout
SetOutputQuality pActiveView, CInt(txtOutputRatio.Text)
Set pMapFrame = pGraphicsContainer.FindFrame(pMxDoc.Maps.Item(0))
Set pClone = pMxApp.Printer
Set pPrinter = pClone.Clone

If optMapPortrait Then
    pPrinter.Paper.Orientation = 1  '1 = portrait, '2 = landscape
    Set pMxApp.Printer = pPrinter
    pPageLayout.Page.Orientation = 1
    pActiveView.Refresh
    AdjustMapFrame 0.6005, 4.2068, (0.6005 + 19.7907), (4.2068 + 24.8727) 'reset map to portrait template size
Else
    pPrinter.Paper.Orientation = 2  '1 = portrait, '2 = landscape
    Set pMxApp.Printer = pPrinter
    pPageLayout.Page.Orientation = 2
    pActiveView.Refresh
    AdjustMapFrame 0.6005, 4.2068, (4.2068 + 24.8727), (0.6005 + 19.7907)
End If
Set pActiveView = pPageLayout
pActiveView.PrinterChanged pPrinter
'Set maplayout scale round number, depending on scale
pMapFrame.ExtentType = esriExtentTypeEnum.esriExtentDefault
pMapFrame.MapScale = 0
pMapFrame.ExtentType = esriExtentTypeEnum.esriExtentScale
CurrentMapScale = pMapFrame.MapScale

If chkRoundOffScale Then
        pMapFrame.MapScale = NewMapScale(CurrentMapScale)
    Else
        pMapFrame.MapScale = CurrentMapScale
End If
pMapFrame.ExtentType = esriExtentTypeEnum.esriExtentScale
pActiveView.Refresh
pMapFrame.ExtentType = esriExtentDefault

'Clean up the current map in preparation for the new output, assume there may be a previous output
'The SEPA title is arial font 22
'If txtMapTitle.Text is longer than 15 characters it must be split into several lines on the space
Dim intActiveLineLength As Integer
Dim strTextForTitle As String
If Len(txtMapTitle.Text) > 20 Then
    Dim splitTitle As Variant
    splitTitle = Split(txtMapTitle.Text, " ")
    strTextForTitle = splitTitle(0)
    intActiveLineLength = Len(strTextForTitle)
    For i = 1 To UBound(splitTitle)
        If (intActiveLineLength + Len(splitTitle(i))) > 20 Then
            strTextForTitle = strTextForTitle & vbCrLf & splitTitle(i)
            intActiveLineLength = Len(splitTitle(i))
        Else
            intActiveLineLength = intActiveLineLength + Len(splitTitle(i)) + 1
            strTextForTitle = strTextForTitle & " " & splitTitle(i)
        End If
    Next
End If
ReplaceTitleText strTextForTitle

'create map output resolution
If Not IsNumeric(txtOutputRes.Text) Then
    MsgBox "Please modify the text for the output resolution of the map - it should contain only a number", vbCritical
    Exit Sub
End If

'test txtOutputReport.text for suitability
'check files for output
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(txtOutputReport.Text) Then
    If fso.FolderExists(txtOutputReport.Text) Then
        MsgBox txtOutputReport.Text & " is an existing folder/path that can't be overwritten." & vbCrLf & _
                "Please enter your output file", vbInformation, "Warning"
        Exit Sub
    End If
End If

pExport.ExportFileName = txtOutputReport
Dim tmpDC As Long
tmpDC = GetDC(0)
iScreenResolution = GetDeviceCaps(tmpDC, 88)   '88 is the win32 const for Logical pixels/inch in X)'96  'default screen resolution is usually 96dpi
ReleaseDC 0, tmpDC
iOutputResolution = CLng(txtOutputRes.Text)
pExport.Resolution = iOutputResolution

With exportRECT
    .Left = 0
    .Top = 0
    .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
    .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
End With

Set pPixelBoundsEnv = New Envelope
pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
pExport.PixelBounds = pPixelBoundsEnv

'define raster settings for raster export
Dim pOutputRasterSettings As IOutputRasterSettings
If TypeOf pExport Is IExportImage Then
  'always set the output quality of the display to 1 for image export formats
    SetOutputQuality pActiveView, 1
ElseIf TypeOf pExport Is IOutputRasterSettings Then
Dim lonMapIndex As Long
Dim pMaps As IMaps
Set pMaps = pMxDoc.Maps
For lonMapIndex = 0 To (pMaps.Count - 1)
    Set pActiveView = pMaps.Item(lonMapIndex)
    SetOutputQuality pActiveView, CInt(txtOutputRatio.Text)
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    Set pOutputRasterSettings = pExport
    pOutputRasterSettings.ResampleRatio = CInt(txtOutputRatio.Text)
    Set pOutputRasterSettings = Nothing
Next
End If
Set pActiveView = pPageLayout

hdc = pExport.StartExporting
pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
pExport.FinishExporting
pExport.Cleanup

If optMapLandscape Then
'revert size and paper to portrait for other exports
    pPrinter.Paper.Orientation = 1  '1 = portrait, '2 = landscape
    Set pMxApp.Printer = pPrinter
    pPageLayout.Page.Orientation = 1
    AdjustMapFrame 0.6005, 4.2068, (0.6005 + 19.7907), (4.2068 + 24.8727) 'reset map to portrait template size
End If

'#################################################################################################################################
'Minimise the map and create a table of TP & J stats and export again to page 2, include the chkProduceResultsCSV here
'#################################################################################################################################
If chkProduceResultsTable Then
    AdjustMapFrame 0, 0, 0, 0  'set map to be effectively invisible
    
'add the table, six columns - GBLAKES_ID, site name, order, TP, J, Ref, type
'one title row and one row per GBLAKES_IDs in the network
'however, maximum of 28 records per page.
Dim lonCountRecordsProcessed As Long
Dim intPageNumber As Integer
Dim intTotalPages As Integer
Dim lonStartRecord As Long
lonStartRecord = 0
intPageNumber = 1
lonCountRecordsProcessed = 0
intTotalPages = (UBound(CatchNetRship, 1) \ 28) + 1
If UBound(CatchNetRship, 1) + 1 > 28 Then
    While lonCountRecordsProcessed < UBound(CatchNetRship, 1) ' + 1
        GenerateTable UBound(CatchNetRship, 1) + 1, lonStartRecord, intPageNumber, intTotalPages
        Dim strRight As String
        Dim strLeft As String
        strRight = Right(txtOutputReport, 3)
        strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
        pExport.ExportFileName = strLeft & "_" & intPageNumber & "_P_table." & strRight
        iScreenResolution = 200  'default screen resolution is usually 96dpi
        iOutputResolution = CLng(txtOutputRes.Text)
        pExport.Resolution = iOutputResolution
        
        With exportRECT
        .Left = 0
        .Top = 0
        .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
        .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
        End With
        
        Set pPixelBoundsEnv = New Envelope
        pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
        pExport.PixelBounds = pPixelBoundsEnv
        
        hdc = pExport.StartExporting
        pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
        pExport.FinishExporting
        pExport.Cleanup
            
        'remove the table from the graphics container
        Dim pGraphicsContainerSelect As IGraphicsContainerSelect
        Set pGraphicsContainerSelect = pGraphicsContainer
        Dim pGraphicsElement As IGraphicElement
        
        Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        While Not pGraphicsElement Is Nothing
            pGraphicsContainer.DeleteElement pGraphicsElement
            Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        Wend
        
        lonStartRecord = lonStartRecord + 28
        lonCountRecordsProcessed = lonCountRecordsProcessed + 28
        intPageNumber = intPageNumber + 1
        
     Wend
Else
    GenerateTable (UBound(CatchNetRship, 1) + 1), 0, 1, 1
    'export it, then modify the file name for the second part
    strRight = Right(txtOutputReport, 3)
    strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
    pExport.ExportFileName = strLeft & "_P_table." & strRight
    iScreenResolution = 200  'default screen resolution is usually 96dpi
    iOutputResolution = CLng(txtOutputRes.Text)
    pExport.Resolution = iOutputResolution
    
    With exportRECT
    .Left = 0
    .Top = 0
    .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
    .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
    End With
    
    Set pPixelBoundsEnv = New Envelope
    pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
    pExport.PixelBounds = pPixelBoundsEnv
    
    hdc = pExport.StartExporting
    pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
    pExport.FinishExporting
    pExport.Cleanup
        
    'remove the table from the graphics container
    Set pGraphicsContainerSelect = pGraphicsContainer
       
    Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
    While Not pGraphicsElement Is Nothing
        pGraphicsContainer.DeleteElement pGraphicsElement
        Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
    Wend
        
End If

'################################################################################
'Now create the measured stats table
'################################################################################
lonStartRecord = 0
intPageNumber = 1
lonCountRecordsProcessed = 0
intTotalPages = (UBound(CatchNetRship, 1) \ 28) + 1
If UBound(CatchNetRship, 1) + 1 > 28 Then
    While lonCountRecordsProcessed < UBound(CatchNetRship, 1) ' + 1
        GenerateTableMeas UBound(CatchNetRship, 1) + 1, lonStartRecord, intPageNumber, intTotalPages
        strRight = Right(txtOutputReport, 3)
        strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
        pExport.ExportFileName = strLeft & "_" & intPageNumber & "_MeasP_table." & strRight
        iScreenResolution = 200  'default screen resolution is usually 96dpi
        iOutputResolution = CLng(txtOutputRes.Text)
        pExport.Resolution = iOutputResolution
        
        With exportRECT
        .Left = 0
        .Top = 0
        .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
        .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
        End With
        
        Set pPixelBoundsEnv = New Envelope
        pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
        pExport.PixelBounds = pPixelBoundsEnv
        
        hdc = pExport.StartExporting
        pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
        pExport.FinishExporting
        pExport.Cleanup
            
        'remove the table from the graphics container
        Set pGraphicsContainerSelect = pGraphicsContainer
        Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        While Not pGraphicsElement Is Nothing
            pGraphicsContainer.DeleteElement pGraphicsElement
            Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        Wend
        
        lonStartRecord = lonStartRecord + 28
        lonCountRecordsProcessed = lonCountRecordsProcessed + 28
        intPageNumber = intPageNumber + 1
        
     Wend
Else
    GenerateTableMeas (UBound(CatchNetRship, 1) + 1), 0, 1, 1
    'export it, modify the file name for the second part
    strRight = Right(txtOutputReport, 3)
    strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
    pExport.ExportFileName = strLeft & "_MeasP_table." & strRight
    iScreenResolution = 200  'default screen resolution is usually 96dpi
    iOutputResolution = CLng(txtOutputRes.Text)
    pExport.Resolution = iOutputResolution
    
    With exportRECT
    .Left = 0
    .Top = 0
    .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
    .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
    End With
    
    Set pPixelBoundsEnv = New Envelope
    pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
    pExport.PixelBounds = pPixelBoundsEnv
    
    hdc = pExport.StartExporting
    pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
    pExport.FinishExporting
    pExport.Cleanup
        
    'remove the table from the graphics container
    Set pGraphicsContainerSelect = pGraphicsContainer
       
    Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
    While Not pGraphicsElement Is Nothing
        pGraphicsContainer.DeleteElement pGraphicsElement
        Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
    Wend
        
End If

'################################################################################
'Now create the Modelled capacity to upgrade and downgrade for selected and connected catchments table for the selected site
'################################################################################
'GB Lakes ID, WB ID,   Name, RAG, Cap. to down TP,  Cap. to down J, Cap. to upgrd TP, Cap. to upgrd J, Modelled capacity to upgrade and downgrade for … and connected catchments
Dim strModMeasd As String
Dim L As Integer
For L = 0 To 1
    If L = 0 Then
        strModMeasd = "Modelled"
    Else
        strModMeasd = "Measured"
    End If
    lonStartRecord = 0
    intPageNumber = 1
    lonCountRecordsProcessed = 0
    If UBound(CatchNetRship, 1) + 1 > 28 Then
        While lonCountRecordsProcessed < UBound(CatchNetRship, 1)
            GenerateTableCapacity UBound(CatchNetRship, 1) + 1, lonStartRecord, intPageNumber, intTotalPages, strModMeasd
            strRight = Right(txtOutputReport, 3)
            strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
            If strModMeasd = "Modelled" Then
                pExport.ExportFileName = strLeft & "_" & intPageNumber & "_ModelledCap_table." & strRight ' txtOutputReport
            Else
                pExport.ExportFileName = strLeft & "_" & intPageNumber & "_MeasuredCap_table." & strRight ' txtOutputReport
            End If
            iScreenResolution = 200  'default screen resolution is usually 96dpi
            iOutputResolution = CLng(txtOutputRes.Text)
            pExport.Resolution = iOutputResolution
            
            With exportRECT
            .Left = 0
            .Top = 0
            .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
            .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
            End With
            
            Set pPixelBoundsEnv = New Envelope
            pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
            pExport.PixelBounds = pPixelBoundsEnv
            
            hdc = pExport.StartExporting
            pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
            pExport.FinishExporting
            pExport.Cleanup
                
            'remove the table from the graphics container
            Set pGraphicsContainerSelect = pGraphicsContainer
            Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
            While Not pGraphicsElement Is Nothing
                pGraphicsContainer.DeleteElement pGraphicsElement
                Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
            Wend
            lonStartRecord = lonStartRecord + 28
            lonCountRecordsProcessed = lonCountRecordsProcessed + 28
            intPageNumber = intPageNumber + 1
        Wend
    Else
        GenerateTableCapacity (UBound(CatchNetRship, 1) + 1), 0, 1, 1, strModMeasd
        'export it, modify the file name for the second part
        strRight = Right(txtOutputReport, 3)
        strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
        If strModMeasd = "Modelled" Then
            pExport.ExportFileName = strLeft & "_ModelledCap_table." & strRight
        Else
             pExport.ExportFileName = strLeft & "_MeasuredCap_table." & strRight
        End If
        iScreenResolution = 200  'default screen resolution is usually 96dpi
        iOutputResolution = CLng(txtOutputRes.Text)
        pExport.Resolution = iOutputResolution
        
        With exportRECT
        .Left = 0
        .Top = 0
        .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
        .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
        End With
        
        Set pPixelBoundsEnv = New Envelope
        pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
        pExport.PixelBounds = pPixelBoundsEnv
        
        hdc = pExport.StartExporting
        pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
        pExport.FinishExporting
        pExport.Cleanup
            
        'remove the table from the graphics container
        Set pGraphicsContainerSelect = pGraphicsContainer
           
        Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        While Not pGraphicsElement Is Nothing
            pGraphicsContainer.DeleteElement pGraphicsElement
            Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
        Wend
    End If
Next

'################################################################################
'Now create the land cover table for the selected site
'################################################################################
'No site has more than 23 different land covers so no requirement for multi-page
Const intNumberOfLCoverColumns = 5
Dim strLCoverArrayColumnHeaders(intNumberOfLCoverColumns) As String
strLCoverArrayColumnHeaders(0) = "P load (kg  per year)"
strLCoverArrayColumnHeaders(1) = "Land cover"
strLCoverArrayColumnHeaders(2) = "Area (m" & Chr(178) & ")"
strLCoverArrayColumnHeaders(3) = "Percent of site (area)"
strLCoverArrayColumnHeaders(4) = "User" & vbCrLf & "modified"
strLCoverArrayColumnHeaders(5) = ""
Dim dblLCoverColumnWidth(intNumberOfLCoverColumns) As Double
dblLCoverColumnWidth(0) = 3.2
dblLCoverColumnWidth(1) = 9.3
dblLCoverColumnWidth(2) = 2.5
dblLCoverColumnWidth(3) = 2.5
dblLCoverColumnWidth(4) = 2.2904
dblLCoverColumnWidth(5) = 0
'add a summary line by using a copy of arrayCatchPforChosenGBLAKES_ID() with an extra row
'sum of land cover for selected site is: Format(dblSumLocalInputs, "#.0")
'sum of area of land cover for selected site is:
'if is modified then sum of arrayCatchPforChosenGBLAKES_ID(i, 5)
'if is not modified then arrayCatchPforChosenGBLAKES_ID(i, 3)

Dim dblAreaSumForReport As Double
dblAreaSumForReport = 0
Dim dblPSumForReport As Double
dblPSumForReport = 0
Dim dblPSumModForReport As Double
dblPSumModForReport = 0
If optReportOnBaseline Then
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
        dblAreaSumForReport = dblAreaSumForReport + arrayCatchPforChosenGBLAKES_ID(i, 3)
        dblPSumForReport = dblPSumForReport + arrayCatchPforChosenGBLAKES_ID(i, 2)
        dblPSumModForReport = dblPSumModForReport + arrayCatchPforChosenGBLAKES_ID(i, 7)
    Next
Else
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
        dblAreaSumForReport = dblAreaSumForReport + arrayCatchPforChosenGBLAKES_ID(i, 5)
        dblPSumModForReport = dblPSumModForReport + arrayCatchPforChosenGBLAKES_ID(i, 7)
    Next
End If
'copy the array to another with an extra row at the bottom for summary data.
Dim intTemp As Integer
intTemp = UBound(arrayCatchPforChosenGBLAKES_ID, 1) + 1
ReDim arrayCatchPforChosenGBLAKES_ID_With_Summary(intTemp, 7)
For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 0) = arrayCatchPforChosenGBLAKES_ID(i, 0)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 1) = arrayCatchPforChosenGBLAKES_ID(i, 1)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 2) = arrayCatchPforChosenGBLAKES_ID(i, 2)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 3) = arrayCatchPforChosenGBLAKES_ID(i, 3)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 4) = arrayCatchPforChosenGBLAKES_ID(i, 4)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 5) = arrayCatchPforChosenGBLAKES_ID(i, 5)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 6) = arrayCatchPforChosenGBLAKES_ID(i, 6)
    arrayCatchPforChosenGBLAKES_ID_With_Summary(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 7)
Next

arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 0) = ""
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 1) = "Total for all land covers"
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 2) = Format(dblSumLocalInputs, "#.0")
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 3) = dblAreaSumForReport
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 4) = ""
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 5) = dblAreaSumForReport
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 6) = ""
arrayCatchPforChosenGBLAKES_ID_With_Summary(UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary), 7) = Format(dblPSumModForReport, "#.0")

'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
GenerateOutputTable (UBound(arrayCatchPforChosenGBLAKES_ID_With_Summary, 1) + 1), 0, 1, 1, intNumberOfLCoverColumns, _
                        strLCoverArrayColumnHeaders(), dblLCoverColumnWidth(), arrayCatchPforChosenGBLAKES_ID_With_Summary(), "LandCover"

'export it, modify the file name for the second part
strRight = Right(txtOutputReport, 3)
strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
pExport.ExportFileName = strLeft & "_LCover_table." & strRight ' txtOutputReport
iScreenResolution = 200  'default screen resolution is usually 96dpi
iOutputResolution = CLng(txtOutputRes.Text)
pExport.Resolution = iOutputResolution

With exportRECT
    .Left = 0
    .Top = 0
    .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
    .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
End With

Set pPixelBoundsEnv = New Envelope
pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
pExport.PixelBounds = pPixelBoundsEnv

hdc = pExport.StartExporting
pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing

pExport.FinishExporting
pExport.Cleanup
    
'remove the table from the graphics container
Set pGraphicsContainerSelect = pGraphicsContainer
   
Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
While Not pGraphicsElement Is Nothing
    pGraphicsContainer.DeleteElement pGraphicsElement
    Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
Wend
    
'################################################################################
'Export summary and sewage loads to page 4
'################################################################################
'create the new name
 strRight = Right(txtOutputReport, 3)
 strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
 pExport.ExportFileName = strLeft & "_Summary_table." & strRight ' txtOutputReport
 iScreenResolution = 200  'default screen resolution is usually 96dpi
 iOutputResolution = CLng(txtOutputRes.Text)
 pExport.Resolution = iOutputResolution
'################################################################################
'generate the new table. This can only be a one line long
'Urban load (kg P), Urban pop, Rural load (kg P), Rural pop
'need to add the point sources into the summary page
'################################################################################
 Const intNumberOfSummaryColumns = 8
 Dim strSewageArrayColumnHeaders(intNumberOfSummaryColumns) As String
 strSewageArrayColumnHeaders(0) = "Sum P" & vbCrLf & "(kg)"
 strSewageArrayColumnHeaders(1) = "Sum land" & vbCrLf & "cover" & vbCrLf & "input" & vbCrLf & "P (kg)"
 strSewageArrayColumnHeaders(2) = "Sum" & vbCrLf & "upstream" & vbCrLf & "input" & vbCrLf & "P (kg)"
 strSewageArrayColumnHeaders(3) = "Urban" & vbCrLf & "pop."
 strSewageArrayColumnHeaders(4) = "Urban P" & vbCrLf & "(kg)"
 strSewageArrayColumnHeaders(5) = "Rural" & vbCrLf & "pop."
 strSewageArrayColumnHeaders(6) = "Rural P" & vbCrLf & "(kg)"
 strSewageArrayColumnHeaders(7) = "Other" & vbCrLf & "point" & vbCrLf & "sources" & vbCrLf & "(kg)"
 strSewageArrayColumnHeaders(8) = ""
 Dim dblSewageColumnWidth(intNumberOfSummaryColumns) As Double
 dblSewageColumnWidth(0) = 2.3
 dblSewageColumnWidth(1) = 2.9
 dblSewageColumnWidth(2) = 2.9
 dblSewageColumnWidth(3) = 2.3
 dblSewageColumnWidth(4) = 2.4904
 dblSewageColumnWidth(5) = 2.3
 dblSewageColumnWidth(6) = 2.4
 dblSewageColumnWidth(7) = 2.2
 dblSewageColumnWidth(8) = 0

'varCatchmentSewage contains the read-in data, CatchNetRship() contains scenario
 GenerateOutputTable 1, 0, 1, 1, intNumberOfSummaryColumns, strSewageArrayColumnHeaders(), _
     dblSewageColumnWidth(), CatchNetRship(), "Summary"

 With exportRECT
 .Left = 0
 .Top = 0
 .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
 .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
 End With
 
 Set pPixelBoundsEnv = New Envelope
 pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
 pExport.PixelBounds = pPixelBoundsEnv
 
 hdc = pExport.StartExporting
 pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing

 pExport.FinishExporting
 pExport.Cleanup
     
 'remove the table from the graphics container
 Set pGraphicsContainerSelect = pGraphicsContainer
    
 Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
 While Not pGraphicsElement Is Nothing
     pGraphicsContainer.DeleteElement pGraphicsElement
     Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
 Wend

'################################################################################
'Export point sources to page 5
'################################################################################
'generate the new table. This is as long as the number of point sources in the network - so could exceed a page
'GBLAKES_ID, point source, amount in point source
 Const intNumberOfPointSourceColumns = 4
 Dim strPointSourcesArrayColumnHeaders(intNumberOfPointSourceColumns) As String
 strPointSourcesArrayColumnHeaders(0) = "GB Lakes ID"
 strPointSourcesArrayColumnHeaders(1) = "WB ID"
 strPointSourcesArrayColumnHeaders(2) = "Point source type"
 strPointSourcesArrayColumnHeaders(3) = "Amount (kg)"
 Dim dblPointSourcesColumnWidth(intNumberOfSummaryColumns) As Double
 dblPointSourcesColumnWidth(0) = 3
 dblPointSourcesColumnWidth(1) = 3
 dblPointSourcesColumnWidth(2) = 9.7968
 dblPointSourcesColumnWidth(3) = 3.9934

intTotalPages = 1
Dim lonRecordsInPointSource As Long
Dim blnArrayIsNotEmpty As Boolean
blnArrayIsNotEmpty = False
On Error Resume Next
blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
If blnArrayIsNotEmpty Then
    intTotalPages = (UBound(varPointSource, 1) \ 28) + 1
    lonRecordsInPointSource = UBound(varPointSource, 1)
Else
    intTotalPages = 1
    lonRecordsInPointSource = 0
End If
Dim lonRowsToProcess As Long
If intTotalPages > 1 Then 'it's a multi-page output
    lonCountRecordsProcessed = 0
    intPageNumber = 1
    While intPageNumber <= intTotalPages
        'create the new name
         strRight = Right(txtOutputReport, 3)
         strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
         pExport.ExportFileName = strLeft & "_" & intPageNumber & "_PointSources_table." & strRight ' txtOutputReport
         iScreenResolution = 200  'default screen resolution is usually 96dpi
         iOutputResolution = CLng(txtOutputRes.Text)
         pExport.Resolution = iOutputResolution
         
         'varPointSource contains the data
        GenerateOutputTable ((lonRecordsInPointSource + 1) - ((intPageNumber - 1) * 28)), ((intPageNumber - 1) * 28), intPageNumber, intTotalPages, _
        intNumberOfPointSourceColumns, strPointSourcesArrayColumnHeaders(), dblPointSourcesColumnWidth(), varPointSource(), "PointSources"
        With exportRECT
         .Left = 0
         .Top = 0
         .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
         .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
         End With
         
         Set pPixelBoundsEnv = New Envelope
         pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
         pExport.PixelBounds = pPixelBoundsEnv
         
         hdc = pExport.StartExporting
         pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
        
         pExport.FinishExporting
         pExport.Cleanup
             
         'remove the table from the graphics container
         Set pGraphicsContainerSelect = pGraphicsContainer
            
         Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
         While Not pGraphicsElement Is Nothing
             pGraphicsContainer.DeleteElement pGraphicsElement
             Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
         Wend
        intPageNumber = intPageNumber + 1
    Wend
Else
    'create the new name
     strRight = Right(txtOutputReport, 3)
     strLeft = Left(txtOutputReport, Len(txtOutputReport) - 4)
     pExport.ExportFileName = strLeft & "_PointSources_table." & strRight ' txtOutputReport
     iScreenResolution = 200  'default screen resolution is usually 96dpi
     iOutputResolution = CLng(txtOutputRes.Text)
     pExport.Resolution = iOutputResolution
    
    'varPointSource contains the data
     GenerateOutputTable (lonRecordsInPointSource + 1), 0, 1, 1, intNumberOfPointSourceColumns, strPointSourcesArrayColumnHeaders(), _
         dblPointSourcesColumnWidth(), varPointSource(), "PointSources"
    
     With exportRECT
     .Left = 0
     .Top = 0
     .Right = pActiveView.exportFrame.Right * (iOutputResolution / iScreenResolution)
     .bottom = pActiveView.exportFrame.bottom * (iOutputResolution / iScreenResolution)
     End With
     
     Set pPixelBoundsEnv = New Envelope
     pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
     pExport.PixelBounds = pPixelBoundsEnv
     
     hdc = pExport.StartExporting
     pActiveView.Output hdc, pExport.Resolution, exportRECT, Nothing, Nothing
    
     pExport.FinishExporting
     pExport.Cleanup
         
     'remove the table from the graphics container
     Set pGraphicsContainerSelect = pGraphicsContainer
        
     Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
     While Not pGraphicsElement Is Nothing
         pGraphicsContainer.DeleteElement pGraphicsElement
         Set pGraphicsElement = pGraphicsContainerSelect.SelectedElements.Next
     Wend
    End If 'this is the end of "If UBound(CatchNetRship, 1) + 1 > 28 Then"
End If  'this is the end of "If chkProduceResultsTable Then"

If chkProduceResultsCSV Then
'################################################################################
'export the six text files
'################################################################################
Open strArrayTextFileNames(0) For Output As #1
Open strArrayTextFileNames(1) For Output As #2
Open strArrayTextFileNames(2) For Output As #3
Open strArrayTextFileNames(3) For Output As #4
Open strArrayTextFileNames(4) For Output As #5
Open strArrayTextFileNames(5) For Output As #6

Print #1, "ID,WBID,SiteName,Relationship,ChosenSite"
Print #2, "ID,WBID,SiteName,ModelledStatus,SEPAStatus,Order,ModelledTP,J,MeasTP,DerivedJ,RefType"
Print #3, "ID,WBID,P,LandCover,Area"
Print #4, "ID,WBID,SumP,SumLCover,SumUpstream,UrbanPop,UrbanP,RuralPop,RuralP"
Print #5, "ID,WBID,PointSource,Amount"
Print #6, "ID,WBID,SiteName,ModelledTP,J,MeasTP,DerivedJ,TP Mod.Cap.to downgrade,J Mod.Cap.to downgrade,TP Mod.Cap.to upgrade,J Mod.Cap.to upgrade,TP Meas.Cap.to downgrade,J Meas.Cap.to downgrade,TP Meas.Cap.to upgrade,J Meas.Cap.to upgrade"

'1. Catchment relationships: Site, Water body ID, Site Name, Relationship to chosen, Chosen site
For j = 0 To UBound(CatchNetRship, 1)
    Print #1, CatchNetRship(j, 0) & "," & ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0))) & "," & ReturnSitename(CatchNetRship(j, 0)) _
    & "," & CatchNetRship(j, 1) & "," & lonChosenGBLAKES_ID
Next

'2. Catchment P: Site, WBID, Site Name, Status, SEPA Status, Order, TP, J, Ref. type - for baseline and scenarios
For j = 0 To UBound(CatchNetRship, 1)
'when a scenario is running (e.g. changed the coefficient of a land cover) I have disabled baseline reporting the values
'below correspond to those in the table on the front tab of the application. This also applies to the JPEG/PDF
        Print #2, CatchNetRship(j, 0) & "," & ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0))) & "," & ReturnSitename(CatchNetRship(j, 0)) _
            & "," & CatchNetRship(j, 21) & "," & ReturnSEPA_Status(CatchNetRship(j, 0)) & "," & lonOrderMatchArray(j) & "," _
            & Format(CatchNetRship(j, 14), "#.0") & "," & Format(CatchNetRship(j, 10), "#.0") & "," & Format(CatchNetRship(j, 25), "#.0") & "," _
            & Format(CatchNetRship(j, 26), "#.0") & "," & CatchNetRship(j, 20)
Next

'3. Land cover P: Site, WBID, P, Land cover, Area
dblSumLocalInputs = 0
Dim strOutput As String
Dim lonRowScroller As Long
Dim dblPToUse As Double
Dim dblAreaToUse As Double
For lonRowScroller = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    strOutput = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0) & "," & ReturnWFD_WB_ID(CLng(arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0))) & ","
    If optUserModified Then
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5)
    Else
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)
    End If
    If dblPToUse < 1 Then
        strOutput = strOutput & "0" & Format(dblPToUse, "#.0") & ","
    Else
        strOutput = strOutput & Format(dblPToUse, "#.0") & ","
    End If
    
    If optUserModified Then
        If (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)) _
        Or (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)) Then
            strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & " (mod.)" & ","
        Else
            strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & ","
        End If
    Else
        strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & ","
    End If
    If dblAreaToUse < 1 Then
         strOutput = strOutput & "0" & Format(dblAreaToUse, "#")
    Else
         strOutput = strOutput & Format(dblAreaToUse, "#")
    End If
    Print #3, strOutput
Next

'4. Summary: Site, WB ID, Sum P, Sum land cover input, Sum upstream input, Urban pop, Urban P, Rural pop, Rural P, PointSrc
strOutput = lonChosenGBLAKES_ID & "," & ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) & ","
If dblJSelectedCatchment < 1 Then
    strOutput = strOutput & "0" & Format(dblJSelectedCatchment, "#.0") & ","
Else
    strOutput = strOutput & Format(dblJSelectedCatchment, "#.0") & ","
End If
If dblSumLocalInputs < 1 Then
    strOutput = strOutput & "0" & Format(dblSumLocalInputs, "#.0") & ","
Else
    strOutput = strOutput & Format(dblSumLocalInputs, "#.0")
End If
If dblSumUpstream < 1 Then
    strOutput = strOutput & "0" & Format(dblSumUpstream, "#.0") & ","
Else
    strOutput = strOutput & Format(dblSumUpstream, "#.0") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = 0 Then
    strOutput = strOutput & "0" & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 18), "#") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 16) < 1 Then
    strOutput = strOutput & "0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0") & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = 0 Then
    strOutput = strOutput & "0" & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 19), "#") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 17) < 1 Then
    strOutput = strOutput & "0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0") & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0") & ","
End If
strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 22), "#.0")
Print #4, strOutput

'5. Point sources: Site,WBID,PointSrc,Amount
blnArrayIsNotEmpty = False
On Error Resume Next
blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
If blnArrayIsNotEmpty Then
    For i = 0 To UBound(varPointSource, 1)
        strOutput = varPointSource(i, 0) & "," & varPointSource(i, 1) & "," & varPointSource(i, 2) & "," & varPointSource(i, 3)
        Print #5, strOutput
    Next
End If

'6. Capacity to upgrade & downgrade. Take these from the list view box
'Print #6, "Site,WBID,SiteName,ModelledTP,J,MeasTP,DerivedJ,TP Mod.Cap.to downgrade,J Mod.Cap.to downgrade,TP Mod.Cap.to upgrade,J Mod.Cap.to upgrade,TP Meas.Cap.to downgrade,J Meas.Cap.to downgrade,TP Meas.Cap.to upgrade,J Meas.Cap.to upgrade"
'Create a ListItem variable.
Dim itmX As ListItem
strOutput = ""
Dim dblMeasTP As Double
Dim dblDerivedJ As Double
dblMeasTP = 0
dblDerivedJ = 0
For i = 1 To lvwCatchmentRelationships2.ListItems.Count
    Set itmX = lvwCatchmentRelationships2.ListItems(i)
    For j = 0 To UBound(CatchNetRship, 1)
        If itmX.SubItems(2) = CatchNetRship(j, 0) Then
            dblMeasTP = CatchNetRship(j, 25)
            dblDerivedJ = CatchNetRship(j, 26)
        End If
    Next
    'need to get the meas P and derived J from CatchNetRship() as it doesn't seem possible to retrieve a tool tip
    strOutput = itmX.SubItems(2) & "," & ReturnWFD_WB_ID(itmX.SubItems(2)) & "," & itmX.SubItems(3) & "," & itmX.SubItems(5) & "," & itmX.SubItems(6) & ","
    strOutput = strOutput & dblMeasTP & "," & dblDerivedJ & "," & itmX.SubItems(8) & "," & itmX.SubItems(9) & ","
    strOutput = strOutput & itmX.SubItems(10) & "," & itmX.SubItems(11) & "," & itmX.SubItems(12) & "," & itmX.SubItems(13) & "," & itmX.SubItems(14) & "," & itmX.SubItems(15)
    Print #6, strOutput
Next

'close the text files
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6

End If 'If chkProduceResultsCSV Then
'switch to layout
AdjustMapFrame 0.6005, 4.2068, (0.6005 + 19.7907), (4.2068 + 24.8727) 'reset map to template size
'and revert back to data frame
Set pMxDoc.ActiveView = pMxDoc.Maps.Item(0)

If strDiscrepancyInClasses <> "" Then
'export a note on the discrepancy in classes between derived TPBreakPoints & SEPA status
    Dim strFileName As String
    'jpg
    strFileName = Left(txtOutputReport.Text, (Len(txtOutputReport.Text) - 4)) & "_StatusDiscrepancy.txt"
    Open strFileName For Append As #1
    Print #1, "There are one or more discrepancies in status between your selected table and the PLUS+ modelled results"
    Print #1, strDiscrepancyInClasses
    Close #1
    'MsgBox "A text file, " & strFileName & ", has been saved to your output folder which describes discrepancies between the SEPA and PLUS+ modelled status"
End If

'clean up
Set fso = Nothing


End Sub
Private Sub cmdCreateReportBatch()
'This is a stripped down version of the report creator, export to pdf and jpeg have been removed as there
'seemed to be a memory leak somewhere that caused a crash after something like 1200 images were exported.
'This works fine now for the whole 8030 catchments.

Set pMxDoc = ThisDocument

'get the status of chkProduceResultsTable and chkProduceResultsCSV and re-instate them before exporting
Dim blnchkProduceResultsTable As Boolean
Dim blnchkProduceResultsCSV As Boolean
blnchkProduceResultsTable = chkProduceResultsTable.Value
blnchkProduceResultsCSV = chkProduceResultsCSV.Value

'check a valid name has been input and it doesn't end in "\"
If Right(txtOutputReport.Text, 1) = "\" Then
    MsgBox "Warning you don't appear to have entered a root file name for your output maps and report. Please do so." _
    & vbCrLf & "Please also make sure you have a root file name for the text output files.", vbCritical
    Exit Sub
End If
If chkProduceResultsCSV Then 'check a valid name has been input and it doesn't end in "\"
    If Right(txtOutputFile.Text, 1) = "\" Then
        MsgBox "Warning you don't appear to have entered a root file name for your text output files. Please do so.", vbCritical
        Exit Sub
    End If
End If

If chkProduceResultsCSV Then
    'ensure the text output file is a text extension
    If Left(Right(txtOutputFile.Text, 4), 1) = "." Then
            txtOutputFile.Text = Left(txtOutputFile.Text, (Len(txtOutputFile.Text) - 4)) & ".txt"
        Else
            txtOutputFile.Text = txtOutputFile.Text & ".txt"
    End If
End If

chkProduceResultsTable.Value = blnchkProduceResultsTable
chkProduceResultsCSV.Value = blnchkProduceResultsCSV

'create the six text output file names
If chkProduceResultsCSV Then
    Dim strArrayTextFileNames(5) As String
    Dim strTempName As String
    strTempName = Left(txtOutputFile.Text, (Len(txtOutputFile.Text) - 4))
    strArrayTextFileNames(0) = strTempName & "_CatchRelationships.txt"
    strArrayTextFileNames(1) = strTempName & "_CatchP.txt"
    strArrayTextFileNames(2) = strTempName & "_LandCoverP.txt"
    strArrayTextFileNames(3) = strTempName & "_Summary.txt"
    strArrayTextFileNames(4) = strTempName & "_PointSource.txt"
    strArrayTextFileNames(5) = strTempName & "_Capacity.txt"
    
    'check the names are suitable
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim j As Integer
    For j = 0 To UBound(strArrayTextFileNames)
        If Not fso.FileExists(strArrayTextFileNames(j)) Then
            If fso.FolderExists(strArrayTextFileNames(j)) Then
                MsgBox strArrayTextFileNames(j) & " is an existing folder/path that can't be overwritten." & vbCrLf & _
                        "Please enter your output file", vbInformation, "Warning"
                Exit Sub
            End If
        End If
    Next
    Set fso = Nothing
End If

chkProduceResultsTable.Value = blnchkProduceResultsTable
chkProduceResultsCSV.Value = blnchkProduceResultsCSV

'test txtOutputReport.text for suitability
'check files for output
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(txtOutputReport.Text) Then
    If fso.FolderExists(txtOutputReport.Text) Then
        MsgBox txtOutputReport.Text & " is an existing folder/path that can't be overwritten." & vbCrLf & _
                "Please enter your output file", vbInformation, "Warning"
        Exit Sub
    End If
End If

Set fso = Nothing


If chkProduceResultsCSV Then
'################################################################################
'export the six text files
'################################################################################
Open strArrayTextFileNames(0) For Output As #1
Open strArrayTextFileNames(1) For Output As #2
Open strArrayTextFileNames(2) For Output As #3
Open strArrayTextFileNames(3) For Output As #4
Open strArrayTextFileNames(4) For Output As #5
Open strArrayTextFileNames(5) For Output As #6

Print #1, "ID,WBID,SiteName,Relationship,ChosenSite"
Print #2, "ID,WBID,SiteName,ModelledStatus,SEPAStatus,Order,ModelledTP,J,MeasTP,DerivedJ,RefType"
Print #3, "ID,WBID,P,LandCover,Area"
Print #4, "ID,WBID,SumP,SumLCover,SumUpstream,UrbanPop,UrbanP,RuralPop,RuralP"
Print #5, "ID,WBID,PointSource,Amount"
Print #6, "ID,WBID,SiteName,ModelledTP,J,MeasTP,DerivedJ,TP Mod.Cap.to downgrade,J Mod.Cap.to downgrade,TP Mod.Cap.to upgrade,J Mod.Cap.to upgrade,TP Meas.Cap.to downgrade,J Meas.Cap.to downgrade,TP Meas.Cap.to upgrade,J Meas.Cap.to upgrade"

'1. Catchment relationships: Site, Water body ID, Site Name, Relationship to chosen, Chosen site
For j = 0 To UBound(CatchNetRship, 1)
    Print #1, CatchNetRship(j, 0) & "," & ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0))) & "," & ReturnSitename(CatchNetRship(j, 0)) _
    & "," & CatchNetRship(j, 1) & "," & lonChosenGBLAKES_ID
Next

'2. Catchment P: Site, WBID, Site Name, Status, SEPA Status, Order, TP, J, Ref. type - for baseline and scenarios
For j = 0 To UBound(CatchNetRship, 1)
'when a scenario is running (e.g. changed the coefficient of a land cover) I have disabled baseline reporting the values
'below correspond to those in the table on the front tab of the application. This also applies to the JPEG/PDF
        Print #2, CatchNetRship(j, 0) & "," & ReturnWFD_WB_ID(CLng(CatchNetRship(j, 0))) & "," & ReturnSitename(CatchNetRship(j, 0)) _
            & "," & CatchNetRship(j, 21) & "," & ReturnSEPA_Status(CatchNetRship(j, 0)) & "," & lonOrderMatchArray(j) & "," _
            & Format(CatchNetRship(j, 14), "#.0") & "," & Format(CatchNetRship(j, 10), "#.0") & "," & Format(CatchNetRship(j, 25), "#.0") & "," _
            & Format(CatchNetRship(j, 26), "#.0") & "," & CatchNetRship(j, 20)
Next

'3. Land cover P: Site, WBID, P, Land cover, Area
dblSumLocalInputs = 0
Dim strOutput As String
Dim lonRowScroller As Long
Dim dblPToUse As Double
Dim dblAreaToUse As Double
For lonRowScroller = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    strOutput = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0) & "," & ReturnWFD_WB_ID(CLng(arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0))) & ","
    If optUserModified Then
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5)
    Else
        dblPToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)
        dblAreaToUse = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)
    End If
    If dblPToUse < 1 Then
        strOutput = strOutput & "0" & Format(dblPToUse, "#.0") & ","
    Else
        strOutput = strOutput & Format(dblPToUse, "#.0") & ","
    End If
    
    If optUserModified Then
        If (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2)) _
        Or (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5) <> arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3)) Then
            strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & " (mod.)" & ","
        Else
            strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & ","
        End If
    Else
        strOutput = strOutput & arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) & ","
    End If
    If dblAreaToUse < 1 Then
         strOutput = strOutput & "0" & Format(dblAreaToUse, "#")
    Else
         strOutput = strOutput & Format(dblAreaToUse, "#")
    End If
    Print #3, strOutput
Next

'4. Summary: Site, WB ID, Sum P, Sum land cover input, Sum upstream input, Urban pop, Urban P, Rural pop, Rural P, PointSrc
strOutput = lonChosenGBLAKES_ID & "," & ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) & ","
If dblJSelectedCatchment < 1 Then
    strOutput = strOutput & "0" & Format(dblJSelectedCatchment, "#.0") & ","
Else
    strOutput = strOutput & Format(dblJSelectedCatchment, "#.0") & ","
End If
If dblSumLocalInputs < 1 Then
    strOutput = strOutput & "0" & Format(dblSumLocalInputs, "#.0") & ","
Else
    strOutput = strOutput & Format(dblSumLocalInputs, "#.0")
End If
If dblSumUpstream < 1 Then
    strOutput = strOutput & "0" & Format(dblSumUpstream, "#.0") & ","
Else
    strOutput = strOutput & Format(dblSumUpstream, "#.0") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 18) = 0 Then
    strOutput = strOutput & "0" & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 18), "#") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 16) < 1 Then
    strOutput = strOutput & "0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0") & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 16), "#.0") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 19) = 0 Then
    strOutput = strOutput & "0" & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 19), "#") & ","
End If
If CatchNetRship(intIndexSelectedGBLAKES_ID, 17) < 1 Then
    strOutput = strOutput & "0" & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0") & ","
Else
    strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 17), "#.0") & ","
End If
strOutput = strOutput & Format(CatchNetRship(intIndexSelectedGBLAKES_ID, 22), "#.0")
Print #4, strOutput

'5. Point sources: Site,WBID,PointSrc,Amount
Dim blnArrayIsNotEmpty As Boolean
Dim i As Integer

blnArrayIsNotEmpty = False
On Error Resume Next
blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
If blnArrayIsNotEmpty Then
    For i = 0 To UBound(varPointSource, 1)
        strOutput = varPointSource(i, 0) & "," & varPointSource(i, 1) & "," & varPointSource(i, 2) & "," & varPointSource(i, 3)
        Print #5, strOutput
    Next
End If

'6. Capacity to upgrade & downgrade. Take these from the list view box
'Print #6, "Site,WBID,SiteName,ModelledTP,J,MeasTP,DerivedJ,TP Mod.Cap.to downgrade,J Mod.Cap.to downgrade,TP Mod.Cap.to upgrade,J Mod.Cap.to upgrade,TP Meas.Cap.to downgrade,J Meas.Cap.to downgrade,TP Meas.Cap.to upgrade,J Meas.Cap.to upgrade"
'Create a ListItem variable.
Dim itmX As ListItem
strOutput = ""
Dim dblMeasTP As Double
Dim dblDerivedJ As Double
dblMeasTP = 0
dblDerivedJ = 0
For i = 1 To lvwCatchmentRelationships2.ListItems.Count
    Set itmX = lvwCatchmentRelationships2.ListItems(i)
    For j = 0 To UBound(CatchNetRship, 1)
        If itmX.SubItems(2) = CatchNetRship(j, 0) Then
            dblMeasTP = CatchNetRship(j, 25)
            dblDerivedJ = CatchNetRship(j, 26)
        End If
    Next
    'need to get the meas P and derived J from CatchNetRship() as it doesn't seem possible to retrieve a tool tip
    strOutput = itmX.SubItems(2) & "," & ReturnWFD_WB_ID(itmX.SubItems(2)) & "," & itmX.SubItems(3) & "," & itmX.SubItems(5) & "," & itmX.SubItems(6) & ","
    strOutput = strOutput & dblMeasTP & "," & dblDerivedJ & "," & itmX.SubItems(8) & "," & itmX.SubItems(9) & ","
    strOutput = strOutput & itmX.SubItems(10) & "," & itmX.SubItems(11) & "," & itmX.SubItems(12) & "," & itmX.SubItems(13) & "," & itmX.SubItems(14) & "," & itmX.SubItems(15)
    Print #6, strOutput
Next

'close the text files
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6

End If 'If chkProduceResultsCSV Then
'switch to layout

If strDiscrepancyInClasses <> "" Then
'export a note on the discrepancy in classes between derived TPBreakPoints & SEPA status
    Dim strFileName As String
    'jpg
    strFileName = Left(txtOutputReport.Text, (Len(txtOutputReport.Text) - 4)) & "_StatusDiscrepancy.txt"
    Open strFileName For Output As #1
    Print #1, "There are one or more discrepancies in status between your selected table and the PLUS+ modelled results"
    Print #1, strDiscrepancyInClasses
    Close #1
    'MsgBox "A text file, " & strFileName & ", has been saved to your output folder which describes discrepancies between the SEPA and PLUS+ modelled status"
End If

'clean up
Set fso = Nothing


End Sub

Private Sub cmdCreateScenario_Click()
'initialise the database
Dim pPropset As IPropertySet
Set pPropset = New PropertySet
pPropset.SetProperty "DATABASE", strListofGDBContainingScenarioTables(0)
pPropset.SetProperty "DATAPROVIDER", "Access Data Source"
Dim pwf As IWorkspaceFactory
Set pwf = New AccessWorkspaceFactory
Dim fws As IFeatureWorkspace
Set fws = pwf.Open(pPropset, 0)

Dim pTable As ITable
Dim pCursor As ICursor
Dim pCursorScenario As ICursor
Dim pRow As IRow
Dim pRowBuff As IRowBuffer
Dim strRegionText As String
Dim blnLCoverFound As Boolean
Dim Response2 As String

'Open the scenario table
Set pTable = fws.OpenTable(strScenarioTbl)

'Check no problems with tables
If Not blnScenarioCanBeSaved Then
    MsgBox "There is a problem with the tables loaded for your scenario." _
    & vbCrLf & "This scenario cannot be saved." _
    & vbCrLf & "Please ensure the following tables are loaded:" & vbCrLf & _
    strScenarioTblCatchmentSewageName & vbCrLf & _
    strScenarioTblCatchPName & vbCrLf & _
    strScenarioTblExportsName & vbCrLf & _
    strScenarioTblLoadPrecursorName & vbCrLf & _
    strScenarioTblPerCapitaTPLoadsName & vbCrLf & _
    strScenarioTblTPBreakPointsName, vbCrLf & _
    strScenarioTblPointSource, vbCrLf & _
    strScenarioTbl, vbCritical
    Exit Sub
End If

'Option to exit before anything is created
If optOriginalData Then
    If blnDataLoadedFromAScenario Then
        Response2 = MsgBox("You have loaded a scenario data set, however as the option to 'Use original data' has been selected" & vbCrLf & _
        "your output scenario will consist of the original source data, not the scenario." & vbCrLf & "Do you want to continue?", vbYesNo)
        If Response2 = vbNo Then
            Exit Sub
        End If
    End If
End If

'Check that the ScenarioID is numeric
If Not IsNumeric(txtScenarioID.Text) Then
    MsgBox "Your scenario ID is not a number. Please enter a new ID number." _
    & vbCrLf & "This scenario will not be saved with this ID number.", vbCritical
    Exit Sub
End If

Dim Response As String
If txtScenarioName.Text = "" Then
    Response = MsgBox("You have not given your scenario a name, if you wish to save without a name please click 'Yes'", vbYesNo)
    If Response = vbNo Then
        Exit Sub
    End If
End If
If optRegionN Then
    strRegionText = "EQ N"
End If
If optRegionSW Then
    strRegionText = "EQ SW"
End If
If optRegionSE Then
    strRegionText = "EQ SE"
End If

If (Not optRegionN And Not optRegionSW And Not optRegionSE) Then
    If txtBoxOtherRegion.Text = "" Then
        Response = MsgBox("No region entered or selected do you wish to save your scenario without one?", vbYesNo)
            If Response = vbNo Then
                Exit Sub
            Else
                'save the scenario without a region
                strRegionText = ""
            End If
    Else
            strRegionText = txtBoxOtherRegion.Text
    End If
End If

'Check that the ScenarioID is new or offer the user the opportunity to choose another
Dim blnScenarioIDAlreadyExists As Boolean
blnScenarioIDAlreadyExists = False
Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
Set pCursor = pScenario.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
While Not pRow Is Nothing
    If pRow.Value(intScenarioIDField) = CLng(txtScenarioID) Then
        blnScenarioIDAlreadyExists = True
    End If
    Set pRow = pCursor.NextRow
Wend

'Process the scenario table data
'Check the ID number doesn't already exist. If it does prompt the user for another. This tool is not a table editing suite
If blnScenarioIDAlreadyExists Then
    MsgBox "A scenario with ID " & txtScenarioID & " already exists. Please enter a new ID number or edit your tables to remove the old scenario." _
    & vbCrLf & "This scenario will not be saved with this ID number.", vbCritical
    Exit Sub
End If

'Scenario table comes from input boxes
Set pRowBuff = pTable.CreateRowBuffer
'Populate the row with values
pRowBuff.Value(intScenarioIDField) = CLng(txtScenarioID)
pRowBuff.Value(intScenarioNameField) = txtScenarioName
pRowBuff.Value(intScenarioCreatorField) = txtScenarioOwner
pRowBuff.Value(intScenarioCreationDateField) = txtScenarioDate
pRowBuff.Value(intScenarioCommentField) = txtScenarioComment
pRowBuff.Value(intScenarioRegionField) = strRegionText
Set pCursor = pTable.Insert(True)
pCursor.InsertRow pRowBuff

'input data comes from arrays
'There are two approaches to scenarios:
'1. Scenarios created by modifying values in the form
'2. Scenarios created by intersecting a user supplied polygon with a user defined network of catchments
'There is also a secondary type where (2) can be further modified with (1)

'to be copied/modifed to the scenario tables:
'1. LocalCatchment_and_Network_S (a shapefile): GBLAKES_ID, Sitename, Order_, Catch_Net (this number identifies which catchments are connected - 0 means no connections)
'2. LoadPrecursor_S (a table): GBLAKES_ID, LochOrder, LochArea, LochMeanDepth, LocalArea, OECDDenominator, OECDExponentDenominator, LocalRunoff
'3. CatchmentSewage_S (a table): GBLAKES_ID, Urb_Rur, Load, Pop,
'4. CatchP_S (a table): GBLAKES_ID, LCOVDESC, P, Area
'5. TPBreakPoints_S (a table): GBLAKES_ID, Reference_Type, HighGood_P, GoodModerate_P, ModeratePoor_P, PoorBad_P
'6. Exports_S (a table): LCOVCODE, SlopeCode, MatchCode, LCOVDESC, Min, Max, Average
'7. PerCapitaTPLoads_S (a table): Urb_Rur, PerCapitaTPLoad
'8. PointSource_S (a table): GBLAKES_ID, PointSource, Amount

'#######################################################################################
'optOriginalData or optModifiedData
'If optOriginalData then all outputs to scenario tables come straight from the baseline input table - this would be a good way to start
'a run to modify land cover. Note that this does not treat an opened scenario as a source, the use modified option does.
'#######################################################################################
If optOriginalData Then
    'write all the matching read in baseline data to scenario table
    '1. write the polygons matching the catchment network to LocalCatchment_and_Network_S
    CopyCatchmentPolygonsToScenario
    
    'get the various Field ID's
    GetFieldIndices
    
    '2. copy the table data matching the catchment network to scenario tables pCatchmentSewageTable_S
    Set pQueryFilt = New QueryFilter
    pQueryFilt.SubFields = "GBLAKES_ID,Urb_Rur,Load,Pop"
    Dim i As Integer
    pQueryFilt.WhereClause = "GBLAKES_ID = "
    For i = 0 To UBound(CatchNetRship, 1)
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & CatchNetRship(i, 0)
        If i < UBound(CatchNetRship, 1) Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " or GBLAKES_ID = "
        End If
    Next
    'the fields required from the input table are: intSewageGBLAKES_IDField, intUrb_RurField, intLoadField, intPopulationField
    Set pCursor = pCatchmentSewageTable.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblCatchmentSewageName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intSewageGBLAKES_IDField)
        pRowBuff.Value(pTable.FindField("Urb_Rur")) = pRow.Value(intUrb_RurField)
        pRowBuff.Value(pTable.FindField("Load")) = pRow.Value(intLoadField)
        pRowBuff.Value(pTable.FindField("Pop")) = pRow.Value(intPopulationField)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
    'option to incorporate user supplied land cover/slope data
    If blnUseModifiedLandCoverSlope Then
    'use the data in the newly created summary.dbf
        Set pQueryFilt = New QueryFilter
        Set pCursor = pTableUserDefinedLCoverSlope_Summary.Search(pQueryFilt, False)
        Set pRow = pCursor.NextRow
        Set pTable = fws.OpenTable(strScenarioTblCatchPName)
        While Not pRow Is Nothing
            Set pRowBuff = pTable.CreateRowBuffer
            pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("GBLAKES_ID"))
            pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("LCOVDESC"))
            pRowBuff.Value(pTable.FindField("P")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("SUM_P"))
            pRowBuff.Value(pTable.FindField("Area")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("SUM_Area"))
            pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
            Set pCursorScenario = pTable.Insert(True)
            pCursorScenario.InsertRow pRowBuff
            Set pRow = pCursor.NextRow
        Wend
    Else
        'pCatchPTable_S
        Set pQueryFilt = New QueryFilter
        pQueryFilt.SubFields = "GBLAKES_ID,LCOVDESC,P,Area"
        pQueryFilt.WhereClause = "Catch_Net = " & lonChosenNetwork
        Set pCursor = pCatchPTable.Table.Search(pQueryFilt, False)
        Set pRow = pCursor.NextRow
        Set pTable = fws.OpenTable(strScenarioTblCatchPName)
        While Not pRow Is Nothing
            Set pRowBuff = pTable.CreateRowBuffer
            pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intCatchP_GBLAKES_IDField)
            pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(intCatchP_LCOVDESCField)
            pRowBuff.Value(pTable.FindField("P")) = pRow.Value(intCatchP_PField)
            pRowBuff.Value(pTable.FindField("Area")) = pRow.Value(intCatchP_AreaField)
            pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
            Set pCursorScenario = pTable.Insert(True)
            pCursorScenario.InsertRow pRowBuff
            Set pRow = pCursor.NextRow
        Wend
    End If
    
    'pExportsTable_S
    'This table is only used to populate the values in cboResolveAreaDifference - that is in approximate scenario creation where
    'the user selects an alternate land cover without using a new polygon.
    Set pQueryFilt = New QueryFilter
    pQueryFilt.SubFields = "LCOVCODE,SlopeCode,MatchCode,LCOVDESC,Min_,Max_,Average"
    Set pCursor = pExportsTable.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblExportsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("LCOVCODE")) = pRow.Value(intLCOVCODE)
        pRowBuff.Value(pTable.FindField("SlopeCode")) = pRow.Value(intSlopeCode)
        pRowBuff.Value(pTable.FindField("MatchCode")) = pRow.Value(intMatchCode)
        pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(intLCOVDESC)
        pRowBuff.Value(pTable.FindField("Min_")) = pRow.Value(intMin)
        pRowBuff.Value(pTable.FindField("Max_")) = pRow.Value(intMax)
        pRowBuff.Value(pTable.FindField("Average")) = pRow.Value(intAverage)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
    'pLoadPrecursorTable_S
    'GBLAKES_ID, LochOrder, LochArea, LochMeanDepth, LocalArea, OECDDenominator, OECDExponentDenominator, LocalRunoff
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = "Catch_Net = " & lonChosenNetwork
    Set pCursor = pLoadPrecursorTable.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblLoadPrecursorName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intPrecursorGBLAKES_IDField)
        pRowBuff.Value(pTable.FindField("LochOrder")) = pRow.Value(intOrderField)
        pRowBuff.Value(pTable.FindField("LochArea")) = pRow.Value(intLochAreaField)
        pRowBuff.Value(pTable.FindField("LochMeanDepth")) = pRow.Value(intLochDepthField)
        pRowBuff.Value(pTable.FindField("LocalArea")) = pRow.Value(intLocalAreaField)
        pRowBuff.Value(pTable.FindField("OECDDenominator")) = pRow.Value(intOECDDenominatorField)
        pRowBuff.Value(pTable.FindField("OECDExponentDenominator")) = pRow.Value(intOECDExponentDenominatorField)
        pRowBuff.Value(pTable.FindField("LocalRunoff")) = pRow.Value(intLocalRunoffField)
        pRowBuff.Value(pTable.FindField("Catch_Net")) = lonChosenNetwork
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
    'pPerCapitaTPLoads_S
    'Urb_Rur, PerCapitaTPLoad, LowerDensity, UpperDensity
    Set pQueryFilt = New QueryFilter
    Set pCursor = pPerCapitaTPLoads.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblPerCapitaTPLoadsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("Urb_Rur")) = pRow.Value(intPerCapitaUrb_RurField)
        pRowBuff.Value(pTable.FindField("PerCapitaTPLoad")) = pRow.Value(intPerCapitaTPLoadField)
        pRowBuff.Value(pTable.FindField("LowerDensity")) = pRow.Value(intLowerDensityField)
        pRowBuff.Value(pTable.FindField("UpperDensity")) = pRow.Value(intUpperDensityField)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
    'pTPBreakPoints_S - these are required for each GBLAKES_ID
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = "GBLAKES_ID = "
    For i = 0 To UBound(CatchNetRship, 1)
        pQueryFilt.WhereClause = pQueryFilt.WhereClause & CatchNetRship(i, 0)
        If i < UBound(CatchNetRship, 1) Then
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & " or GBLAKES_ID = "
        End If
    Next
    'the fields required from the input table are: GBLAKES_ID, Reference_Type, HighGood_P, GoodModerate_P, ModeratePoor_P, PoorBad_P
    Set pCursor = pTPBreakPoints.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblTPBreakPointsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intGBLAKES_IDFieldBreakPoints)
        pRowBuff.Value(pTable.FindField("Reference_Type")) = pRow.Value(intReference_TypeField)
        pRowBuff.Value(pTable.FindField("HighGood_P")) = pRow.Value(intHighGood_PField)
        pRowBuff.Value(pTable.FindField("GoodModerate_P")) = pRow.Value(intGoodModerate_PField)
        pRowBuff.Value(pTable.FindField("ModeratePoor_P")) = pRow.Value(intModeratePoor_PField)
        pRowBuff.Value(pTable.FindField("PoorBad_P")) = pRow.Value(intPoorBad_PField)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    'write the PointSource_S data - this is strictly baseline, non-scenario, unmodified so these must all be zero and null text
    Set pTable = fws.OpenTable(strScenarioTblPointSource)
    Dim blnArrayIsNotEmpty As Boolean
    blnArrayIsNotEmpty = False
    On Error Resume Next
    blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
    If blnArrayIsNotEmpty Then
        For i = 0 To UBound(varPointSource, 1)
            Set pRowBuff = pTable.CreateRowBuffer
            pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = varPointSource(i, 0)
            pRowBuff.Value(pTable.FindField("PointSource")) = "" 'varPointSource(i, 1)
            pRowBuff.Value(pTable.FindField("Amount")) = 0 'varPointSource(i, 2)
            pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
            Set pCursorScenario = pTable.Insert(True)
            pCursorScenario.InsertRow pRowBuff
        Next
    End If

End If 'end of if optOriginalData then

If optModifiedData Then
'#######################################################################################
'Write the modified data to the scenario tables - a modified version of the above
'#######################################################################################
    'write all the matching read in OR modified data to scenario table
    'write the polygons matching the catchment network to LocalCatchment_and_Network_S
    CopyCatchmentPolygonsToScenario
    
    'get the various Field ID's
    GetFieldIndices
    
    'copy the table data matching the catchment network to scenario tables
    'CatchNetRship(i, 0) = GBLAKES_ID
    'CatchNetRship(i, 16) - Urban Load
    'CatchNetRship(i, 17) - Rural Load
    'CatchNetRship(i, 18) - Urban Pop
    'CatchNetRship(i, 19) - Rural Pop
    'pCatchmentSewageTable_S
    Set pTable = fws.OpenTable(strScenarioTblCatchmentSewageName)
    For i = 0 To UBound(CatchNetRship, 1)
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = CatchNetRship(i, 0)
        'input data may have only rural, but user may wish to add urban too, so incorporate both
        If CatchNetRship(i, 16) > 0 Then
            pRowBuff.Value(pTable.FindField("Urb_Rur")) = "Urban"
            pRowBuff.Value(pTable.FindField("Load")) = CatchNetRship(i, 16)
            pRowBuff.Value(pTable.FindField("Pop")) = CatchNetRship(i, 18)
            pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
            Set pCursorScenario = pTable.Insert(True)
            pCursorScenario.InsertRow pRowBuff
        End If
        'always put in a rural - even if 0 pop & load
        pRowBuff.Value(pTable.FindField("Urb_Rur")) = "Rural"
        pRowBuff.Value(pTable.FindField("Load")) = CatchNetRship(i, 17)
        pRowBuff.Value(pTable.FindField("Pop")) = CatchNetRship(i, 19)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
    Next
    
    'option to incorporate user supplied land cover/slope data
    If blnUseModifiedLandCoverSlope Then
    'use the data in the newly created summary.dbf
        Set pQueryFilt = New QueryFilter
        Set pCursor = pTableUserDefinedLCoverSlope_Summary.Search(pQueryFilt, False)
        Set pRow = pCursor.NextRow
        Set pTable = fws.OpenTable(strScenarioTblCatchPName)
        While Not pRow Is Nothing
            Set pRowBuff = pTable.CreateRowBuffer
            pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("GBLAKES_ID"))
            pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("LCOVDESC"))
            pRowBuff.Value(pTable.FindField("P")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("SUM_P"))
            pRowBuff.Value(pTable.FindField("Area")) = pRow.Value(pTableUserDefinedLCoverSlope_Summary.FindField("SUM_Area"))
            pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
            Set pCursorScenario = pTable.Insert(True)
            pCursorScenario.InsertRow pRowBuff
            Set pRow = pCursor.NextRow
        Wend
    Else
        Set pQueryFilt = New QueryFilter
        pQueryFilt.SubFields = "GBLAKES_ID,LCOVDESC,P,Area"
        'take account of whether or not the source data is a scenario
        If Not blnDataLoadedFromAScenario Then
            pQueryFilt.WhereClause = "Catch_Net = " & lonChosenNetwork
            Set pCursor = pCatchPTable.Table.Search(pQueryFilt, False)
            Set pRow = pCursor.NextRow
            Set pTable = fws.OpenTable(strScenarioTblCatchPName)
            While Not pRow Is Nothing
                Set pRowBuff = pTable.CreateRowBuffer
                pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intCatchP_GBLAKES_IDField)
                pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(intCatchP_LCOVDESCField)
                If chkChangePforNetwork.Value = True And optUserModified Then 'change all catchments, added "And optUserModified "
                    'get the values from arrayCatchPforChosenGBLAKES_ID(i,j)
                    blnLCoverFound = False
                    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
                        If pRow.Value(intCatchP_LCOVDESCField) = arrayCatchPforChosenGBLAKES_ID(i, 1) Then
                            blnLCoverFound = True
                            pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 6) * pRow.Value(intCatchP_AreaField) / 10000
                        End If
                        If Not blnLCoverFound Then 'no match for this land cover found in arrayCatchPforChosenGBLAKES_ID(i,j)
                            pRowBuff.Value(pTable.FindField("P")) = pRow.Value(intCatchP_PField)
                        End If
                    Next
                    If chkChangeArea.Value = True Then
                        If pRow.Value(intCatchP_GBLAKES_IDField) = lonChosenGBLAKES_ID Then
                            pRowBuff.Value(pTable.FindField("P")) = CDbl(txtEnterNewP) * CDbl(txtEnterNewArea) / 10000
                        End If
                    End If
                Else 'no global change
                    blnLCoverFound = False
                    If pRow.Value(intCatchP_GBLAKES_IDField) = lonChosenGBLAKES_ID Then
                    'change for lonchosenGBLAKES_ID only
                    'must treat the current user modified land cover separately, the others get the read-in values
                        For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
                            If pRow.Value(intCatchP_LCOVDESCField) = arrayCatchPforChosenGBLAKES_ID(i, 1) Then
                                'the user will be expecting to have the values he sees in lvwCatchmentInfo saved as the scenario, so use them
                                'As these values are calculated in CalculateTP (which is run after cmdModCatchmentInputs is clicked) the outputs from CalculateTP are used here
                                'use the modified values in arrayCatchPforChosenGBLAKES_ID if present (i.e. if modified)
                                If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Then 'check the area to detect modification
                                    pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 7)
                                    pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 5)
                                Else    'it's not modified
                                    pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 2)
                                    pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 3)
                                End If
                            End If
                        Next
                    Else 'it is not the chosen GBLAKES_ID so there is no modified value, so use the input
                        pRowBuff.Value(pTable.FindField("P")) = pRow.Value(intCatchP_PField)
                        pRowBuff.Value(pTable.FindField("Area")) = pRow.Value(intCatchP_AreaField)
                    End If
                End If
                pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
                Set pCursorScenario = pTable.Insert(True)
                pCursorScenario.InsertRow pRowBuff
                Set pRow = pCursor.NextRow
            Wend
        'Above we are only adding land covers from the read-in table, must also get the user added landcovers
        'from cboResolveAreaDifference. If field 3 = 0 then it must be user added
        'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
        For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
            If arrayCatchPforChosenGBLAKES_ID(i, 3) = 0 Then
                pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = arrayCatchPforChosenGBLAKES_ID(i, 0)
                pRowBuff.Value(pTable.FindField("LCOVDESC")) = arrayCatchPforChosenGBLAKES_ID(i, 1)
                pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 7)
                pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 5)
                Set pCursorScenario = pTable.Insert(True)
                pCursorScenario.InsertRow pRowBuff
            End If
        Next
        End If
        
        If blnDataLoadedFromAScenario Then     'it has been loaded from a scenario - this is the same structure as the above
            pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
            Set pCursor = pCatchPTable_S.Table.Search(pQueryFilt, False)
            Set pRow = pCursor.NextRow
            Set pTable = fws.OpenTable(strScenarioTblCatchPName)    'there are now two pointers to the same table, reading and inserting
            While Not pRow Is Nothing
                Set pRowBuff = pTable.CreateRowBuffer
                pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(pCatchPTable_S.Table.FindField("GBLAKES_ID"))
                pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(pCatchPTable_S.Table.FindField("LCOVDESC"))
                If chkChangePforNetwork.Value = True And optUserModified Then 'change all catchments, added "And optUserModified "
                    'get the values from arrayCatchPforChosenGBLAKES_ID(i,j)
                    blnLCoverFound = False
                    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
                        If pRow.Value(pCatchPTable_S.Table.FindField("LCOVDESC")) = arrayCatchPforChosenGBLAKES_ID(i, 1) Then
                            blnLCoverFound = True
                            pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 6) * pRow.Value(pCatchPTable_S.Table.FindField("P")) / 10000
                        End If
                        If Not blnLCoverFound Then 'no match for this land cover found in arrayCatchPforChosenGBLAKES_ID(i,j)
                            pRowBuff.Value(pTable.FindField("P")) = pRow.Value(pCatchPTable_S.Table.FindField("P"))
                        End If
                    Next
                    If chkChangeArea.Value = True Then
                        If pRow.Value(pCatchPTable_S.Table.FindField("GBLAKES_ID")) = lonChosenGBLAKES_ID Then
                            pRowBuff.Value(pTable.FindField("P")) = CDbl(txtEnterNewP) * CDbl(txtEnterNewArea) / 10000
                        End If
                    End If
                Else 'no global change
                    blnLCoverFound = False
                    If pRow.Value(pCatchPTable_S.Table.FindField("GBLAKES_ID")) = lonChosenGBLAKES_ID Then
                    'change for lonchosenGBLAKES_ID only
                    'must treat the current user modified land cover separately, the others get the read-in values
                        For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
                            If pRow.Value(pCatchPTable_S.Table.FindField("LCOVDESC")) = arrayCatchPforChosenGBLAKES_ID(i, 1) Then
                                'the user is expecting to have the values he sees in lvwCatchmentInfo saved as the scenario, so use them
                                'as these values are calculated in CalculateTP (which is run after cmdModCatchmentInputs is clicked) the outputs from CalculateTP are used here
                                'use the modified values in arrayCatchPforChosenGBLAKES_ID if present (i.e. if modified)
                                If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Then
                                    pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 7)
                                    pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 5)
                                Else    'it's not modified
                                    pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 2)
                                    pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 3)
                                End If
                            End If
                        Next
                    Else 'it is not the chosen GBLAKES_ID so there is no modified value, so use the input
                        pRowBuff.Value(pTable.FindField("P")) = pRow.Value(pCatchPTable_S.Table.FindField("P"))
                        pRowBuff.Value(pTable.FindField("Area")) = pRow.Value(pCatchPTable_S.Table.FindField("Area"))
                    End If
                End If
                pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
                If pRowBuff.Value(pTable.FindField("P")) > 0 Or pRowBuff.Value(pTable.FindField("Area")) > 0 Then
                    Set pCursorScenario = pTable.Insert(True)
                    pCursorScenario.InsertRow pRowBuff
                End If
                Set pRow = pCursor.NextRow
            Wend
        'Above we are only adding land covers from the read-in table, must also get the user added landcovers
        'from cboResolveAreaDifference. If field 3 = 0 then it must be user added
        'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
        For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
            If arrayCatchPforChosenGBLAKES_ID(i, 3) = 0 Then
                If arrayCatchPforChosenGBLAKES_ID(i, 7) > 0 Or arrayCatchPforChosenGBLAKES_ID(i, 5) > 0 Then
                    pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = arrayCatchPforChosenGBLAKES_ID(i, 0)
                    pRowBuff.Value(pTable.FindField("LCOVDESC")) = arrayCatchPforChosenGBLAKES_ID(i, 1)
                    pRowBuff.Value(pTable.FindField("P")) = arrayCatchPforChosenGBLAKES_ID(i, 7)
                    pRowBuff.Value(pTable.FindField("Area")) = arrayCatchPforChosenGBLAKES_ID(i, 5)
                    Set pCursorScenario = pTable.Insert(True)
                    pCursorScenario.InsertRow pRowBuff
                End If
            End If
        Next
        End If
    End If
    
    'pExportsTable_S
    'This table is only used to populate the values in cboResolveAreaDifference - that is in approximate scenario creation where
    'the user selects an alternate land cover without using a new polygon.
    'Save the modified value only when the user has ticked the 'Change P for whole network' box - chkChangePforNetwork, use txtEnterNewP
    Set pQueryFilt = New QueryFilter
    If blnDataLoadedFromAScenario Then
        pQueryFilt.SubFields = "LCOVCODE,SlopeCode,MatchCode,LCOVDESC,Min_,Max_,Average"
        pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
        Set pCursor = pExportsTable_S.Table.Search(pQueryFilt, False)
        intLCOVCODE = pExportsTable_S.Table.FindField("LCOVCODE")
        intSlopeCode = pExportsTable_S.Table.FindField("SlopeCode")
        intMatchCode = pExportsTable_S.Table.FindField("MatchCode")
        intLCOVDESC = pExportsTable_S.Table.FindField("LCOVDESC")
        intMin = pExportsTable_S.Table.FindField("Min_")
        intMax = pExportsTable_S.Table.FindField("Max_")
        intAverage = pExportsTable_S.Table.FindField("Average")
    Else
        pQueryFilt.SubFields = "LCOVCODE,SlopeCode,MatchCode,LCOVDESC,Min_,Max_,Average"
        Set pCursor = pExportsTable.Table.Search(pQueryFilt, False)
        intLCOVCODE = pExportsTable.Table.FindField("LCOVCODE")
        intSlopeCode = pExportsTable.Table.FindField("SlopeCode")
        intMatchCode = pExportsTable.Table.FindField("MatchCode")
        intLCOVDESC = pExportsTable.Table.FindField("LCOVDESC")
        intMin = pExportsTable.Table.FindField("Min_")
        intMax = pExportsTable.Table.FindField("Max_")
        intAverage = pExportsTable.Table.FindField("Average")
    End If
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblExportsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("LCOVCODE")) = pRow.Value(intLCOVCODE)
        pRowBuff.Value(pTable.FindField("SlopeCode")) = pRow.Value(intSlopeCode)
        pRowBuff.Value(pTable.FindField("MatchCode")) = pRow.Value(intMatchCode)
        pRowBuff.Value(pTable.FindField("LCOVDESC")) = pRow.Value(intLCOVDESC)
        'implement option for user entered global value
        If pRow.Value(intLCOVDESC) = strLcoverForNetworkChange Then
            If chkChangePforNetwork.Value = True And optUserModified Then
                pRowBuff.Value(pTable.FindField("Min_")) = 0
                pRowBuff.Value(pTable.FindField("Max_")) = 0
                pRowBuff.Value(pTable.FindField("Average")) = CDbl(txtEnterNewP)
            Else
                pRowBuff.Value(pTable.FindField("Min_")) = pRow.Value(intMin)
                pRowBuff.Value(pTable.FindField("Max_")) = pRow.Value(intMax)
                pRowBuff.Value(pTable.FindField("Average")) = pRow.Value(intAverage)
            End If
        Else
            pRowBuff.Value(pTable.FindField("Min_")) = pRow.Value(intMin)
            pRowBuff.Value(pTable.FindField("Max_")) = pRow.Value(intMax)
            pRowBuff.Value(pTable.FindField("Average")) = pRow.Value(intAverage)
        End If
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
    'pLoadPrecursorTable_S - no scenario options, so this section is identical to above
    'GBLAKES_ID, LochOrder, LochArea, LochMeanDepth, LocalArea, OECDDenominator, OECDExponentDenominator, LocalRunoff
    Set pQueryFilt = New QueryFilter
    If blnDataLoadedFromAScenario Then
        pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
        Set pCursor = pLoadPrecursorTable_S.Table.Search(pQueryFilt, False)
        intPrecursorGBLAKES_IDField = pLoadPrecursorTable_S.Table.FindField("GBLAKES_ID")
        intOrderField = pLoadPrecursorTable_S.Table.FindField("LochOrder")
        intLochAreaField = pLoadPrecursorTable_S.Table.FindField("LochArea")
        intLochDepthField = pLoadPrecursorTable_S.Table.FindField("LochMeanDepth")
        intLocalAreaField = pLoadPrecursorTable_S.Table.FindField("LocalArea")
        intOECDDenominatorField = pLoadPrecursorTable_S.Table.FindField("OECDDenominator")
        intOECDExponentDenominatorField = pLoadPrecursorTable_S.Table.FindField("OECDExponentDenominator")
        intLocalRunoffField = pLoadPrecursorTable_S.Table.FindField("LocalRunoff")
    Else
        pQueryFilt.WhereClause = "Catch_Net = " & lonChosenNetwork
        Set pCursor = pLoadPrecursorTable.Table.Search(pQueryFilt, False)
        intPrecursorGBLAKES_IDField = pLoadPrecursorTable.Table.FindField("GBLAKES_ID")
        intOrderField = pLoadPrecursorTable.Table.FindField("LochOrder")
        intLochAreaField = pLoadPrecursorTable.Table.FindField("LochArea")
        intLochDepthField = pLoadPrecursorTable.Table.FindField("LochMeanDepth")
        intLocalAreaField = pLoadPrecursorTable.Table.FindField("LocalArea")
        intOECDDenominatorField = pLoadPrecursorTable.Table.FindField("OECDDenominator")
        intOECDExponentDenominatorField = pLoadPrecursorTable.Table.FindField("OECDExponentDenominator")
        intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff")
    End If
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblLoadPrecursorName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intPrecursorGBLAKES_IDField)
        pRowBuff.Value(pTable.FindField("LochOrder")) = pRow.Value(intOrderField)
        pRowBuff.Value(pTable.FindField("LochArea")) = pRow.Value(intLochAreaField)
        pRowBuff.Value(pTable.FindField("LochMeanDepth")) = pRow.Value(intLochDepthField)
        pRowBuff.Value(pTable.FindField("LocalArea")) = pRow.Value(intLocalAreaField)
        pRowBuff.Value(pTable.FindField("OECDDenominator")) = pRow.Value(intOECDDenominatorField)
        pRowBuff.Value(pTable.FindField("OECDExponentDenominator")) = pRow.Value(intOECDExponentDenominatorField)
        pRowBuff.Value(pTable.FindField("LocalRunoff")) = pRow.Value(intLocalRunoffField)
        pRowBuff.Value(pTable.FindField("Catch_Net")) = lonChosenNetwork
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend

    'pPerCapitaTPLoads_S
    'Urb_Rur, PerCapitaTPLoad, LowerDensity, UpperDensity
    'with scenario saving want the option for different PerCapitaTPLoad from txtPerCapitaTPLoadUrban and txtPerCapitaTPLoadRural
    Set pQueryFilt = New QueryFilter
    If blnDataLoadedFromAScenario Then
        pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
        Set pCursor = pPerCapitaTPLoads_S.Table.Search(pQueryFilt, False)
        intPerCapitaUrb_RurField = pPerCapitaTPLoads_S.Table.FindField("Urb_Rur")
        intPerCapitaTPLoadField = pPerCapitaTPLoads_S.Table.FindField("PerCapitaTPLoad")
        intLowerDensityField = pPerCapitaTPLoads_S.Table.FindField("LowerDensity")
        intUpperDensityField = pPerCapitaTPLoads_S.Table.FindField("UpperDensity")
    Else
        Set pCursor = pPerCapitaTPLoads.Table.Search(pQueryFilt, False)
        intPerCapitaUrb_RurField = pPerCapitaTPLoads.Table.FindField("Urb_Rur")
        intPerCapitaTPLoadField = pPerCapitaTPLoads.Table.FindField("PerCapitaTPLoad")
        intLowerDensityField = pPerCapitaTPLoads.Table.FindField("LowerDensity")
        intUpperDensityField = pPerCapitaTPLoads.Table.FindField("UpperDensity")
    End If
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblPerCapitaTPLoadsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("Urb_Rur")) = pRow.Value(intPerCapitaUrb_RurField)
        If pRow.Value(intPerCapitaUrb_RurField) = "Urban" Then
            If chkPerCapitaTPLoadUrban.Value = True Then
                pRowBuff.Value(pTable.FindField("PerCapitaTPLoad")) = CDbl(txtPerCapitaTPLoadUrban)
            Else
                pRowBuff.Value(pTable.FindField("PerCapitaTPLoad")) = pRow.Value(intPerCapitaTPLoadField)
            End If
        Else    'it's rural
            If chkPerCapitaTPLoadRural.Value = True Then
                pRowBuff.Value(pTable.FindField("PerCapitaTPLoad")) = CDbl(txtPerCapitaTPLoadRural)
            Else
                pRowBuff.Value(pTable.FindField("PerCapitaTPLoad")) = pRow.Value(intPerCapitaTPLoadField)
            End If
        End If
        pRowBuff.Value(pTable.FindField("LowerDensity")) = pRow.Value(intLowerDensityField)
        pRowBuff.Value(pTable.FindField("UpperDensity")) = pRow.Value(intUpperDensityField)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend

    'pTPBreakPoints_S - these are required for each GBLAKES_ID
    'this is not modified scenario dependent, the code is identical to the above
    Set pQueryFilt = New QueryFilter
    If blnDataLoadedFromAScenario Then
        pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
        Set pCursor = pTPBreakpoints_S.Table.Search(pQueryFilt, False)
        intGBLAKES_IDFieldBreakPoints = pTPBreakpoints_S.Table.FindField("GBLAKES_ID")
        intReference_TypeField = pTPBreakpoints_S.Table.FindField("Reference_Type")
        intHighGood_PField = pTPBreakpoints_S.Table.FindField("HighGood_P")
        intGoodModerate_PField = pTPBreakpoints_S.Table.FindField("GoodModerate_P")
        intModeratePoor_PField = pTPBreakpoints_S.Table.FindField("ModeratePoor_P")
        intPoorBad_PField = pTPBreakpoints_S.Table.FindField("PoorBad_P")
    Else
        pQueryFilt.WhereClause = "GBLAKES_ID = "
        For i = 0 To UBound(CatchNetRship, 1)
            pQueryFilt.WhereClause = pQueryFilt.WhereClause & CatchNetRship(i, 0)
            If i < UBound(CatchNetRship, 1) Then
                pQueryFilt.WhereClause = pQueryFilt.WhereClause & " or GBLAKES_ID = "
            End If
        Next
        'the fields required from the input table are: GBLAKES_ID, Reference_Type, HighGood_P, GoodModerate_P, ModeratePoor_P, PoorBad_P
        Set pCursor = pTPBreakPoints.Table.Search(pQueryFilt, False)
        intGBLAKES_IDFieldBreakPoints = pTPBreakPoints.Table.FindField("GBLAKES_ID")
        intReference_TypeField = pTPBreakPoints.Table.FindField("Reference_Type")
        intHighGood_PField = pTPBreakPoints.Table.FindField("HighGood_P")
        intGoodModerate_PField = pTPBreakPoints.Table.FindField("GoodModerate_P")
        intModeratePoor_PField = pTPBreakPoints.Table.FindField("ModeratePoor_P")
        intPoorBad_PField = pTPBreakPoints.Table.FindField("PoorBad_P")
    End If
    Set pRow = pCursor.NextRow
    Set pTable = fws.OpenTable(strScenarioTblTPBreakPointsName)
    While Not pRow Is Nothing
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = pRow.Value(intGBLAKES_IDFieldBreakPoints)
        pRowBuff.Value(pTable.FindField("Reference_Type")) = pRow.Value(intReference_TypeField)
        pRowBuff.Value(pTable.FindField("HighGood_P")) = pRow.Value(intHighGood_PField)
        pRowBuff.Value(pTable.FindField("GoodModerate_P")) = pRow.Value(intGoodModerate_PField)
        pRowBuff.Value(pTable.FindField("ModeratePoor_P")) = pRow.Value(intModeratePoor_PField)
        pRowBuff.Value(pTable.FindField("PoorBad_P")) = pRow.Value(intPoorBad_PField)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
        Set pRow = pCursor.NextRow
    Wend
    
'write the PointSource_S data - this is strictly baseline, non-scenario, unmodified so these must all be zero and null text
    blnArrayIsNotEmpty = False 'error trapping added 16 Jan 2012
    On Error Resume Next
    blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
    Set pTable = fws.OpenTable(strScenarioTblPointSource)
    If blnArrayIsNotEmpty Then
    For i = 0 To UBound(varPointSource, 1)
        Set pRowBuff = pTable.CreateRowBuffer
        pRowBuff.Value(pTable.FindField("GBLAKES_ID")) = varPointSource(i, 0)
        pRowBuff.Value(pTable.FindField("PointSource")) = varPointSource(i, 1)
        pRowBuff.Value(pTable.FindField("Amount")) = varPointSource(i, 2)
        pRowBuff.Value(pTable.FindField("ScenarioID")) = CLng(txtScenarioID.Text)
        Set pCursorScenario = pTable.Insert(True)
        pCursorScenario.InsertRow pRowBuff
    Next
    End If
End If 'end of "If optModifiedData Then"

End Sub
Private Sub cmdEnlarge_Click()
frmCatch.Height = 580
frmCatch.Width = 698.25
DoEvents
End Sub
Private Sub cmdReduce_Click()
frmCatch.Height = 40
frmCatch.Width = 120
DoEvents
End Sub
Private Sub cmdGetCatchmentInfo_Click()
'GetFieldIndices - this is already in CalcTP
blnSelectedIsOrderZero = False
blnModifySewageLoad = False
blnModifyOtherPointSourceLoad = False
chkPerCapitaTPLoadUrban.Value = False
chkPerCapitaTPLoadRural.Value = False
chkPerCapitaTPLoadUrbanAll.Value = False
chkPerCapitaTPLoadRuralAll.Value = False
chkUrbanPop.Value = False
chkRuralPop.Value = False
Erase varPointSource
'get the chosen catchment GBLAKES_ID
Dim varSplit As Variant
Dim i As Long

If cboGBLAKES_IDs.Text <> "" Then
    varSplit = Split(cboGBLAKES_IDs.Text, " - ")
    If varSplit(0) = cboGBLAKES_IDs.Text Then
        MsgBox "You have chosen a code that does not appear in the data.", vbCritical
        Exit Sub
    End If
    lonChosenGBLAKES_ID = varSplit(0)
    strChosenSitename = varSplit(1)
End If

'create an array of GBLAKES_IDs with that network
Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
If blnScenarioLoaded Then
    pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario 'was cboScenarioID.Value
    lonNumGBLAKES_IDs = pScenarioLocalCatchmentAndNetworkTable.RowCount(pQueryFilt)
Else
    lonNumGBLAKES_IDs = pTableCatchment.RowCount(pQueryFilt)
End If

For i = 0 To lonNumGBLAKES_IDs - 1
'lonGBLAKES_IDArray, strSitenameArray, intOrderArray, lonNetworkArray are populated with pTableCatchment in cmdGetCatchmentInfo
'the same is done with pScenarioLocalCatchmentAndNetworkTable in cmdLoadScenario
    If (lonGBLAKES_IDArray(i) = lonChosenGBLAKES_ID) Then
        lonChosenNetwork = lonNetworkArray(i)
    End If
Next
If lonChosenNetwork = 0 Then
    Label4.Caption = "The chosen catchment has no connections."
    DoEvents
    intMatchingGBLAKES_IDs = 0
    ReDim lonGBLAKES_IDNetworkMatchArray(intMatchingGBLAKES_IDs)
    ReDim lonOrderMatchArray(intMatchingGBLAKES_IDs)
    Erase CatchNetRship()
    ReDim CatchNetRship(0, 26)
    CatchNetRship(0, 0) = lonChosenGBLAKES_ID
    cmdCalcTP.Caption = "Calculate total P concentration (for " & lonChosenGBLAKES_ID & ")"
    cmdCalcTP.ControlTipText = "Calculate total P concentration or load (for " & lonChosenGBLAKES_ID & ")"
    lvwCatchmentRelationships2.Height = 0
    'added to give access to the listview for standalone catchments
    CatchNetRship(0, 0) = lonChosenGBLAKES_ID
    CatchNetRship(0, 1) = "Chosen GBLAKES_ID"
    ReDim lonGBLAKES_IDNetworkMatchArray(0)
    ReDim lonOrderMatchArray(0)
    PopulateListViewCatchmentRelationships
Else
    'step through the GBLAKES_ID array appending to an output array any with the correct network id
    intMatchingGBLAKES_IDs = 0
    ReDim Preserve lonGBLAKES_IDNetworkMatchArray(intMatchingGBLAKES_IDs)
    ReDim Preserve lonOrderMatchArray(intMatchingGBLAKES_IDs)
    For i = 0 To lonNumGBLAKES_IDs - 1
        If (lonNetworkArray(i) = lonChosenNetwork) Then
        'add it to the array
            lonGBLAKES_IDNetworkMatchArray(intMatchingGBLAKES_IDs) = lonGBLAKES_IDArray(i)
            lonOrderMatchArray(intMatchingGBLAKES_IDs) = intOrderArray(i)
            intMatchingGBLAKES_IDs = intMatchingGBLAKES_IDs + 1
            ReDim Preserve lonGBLAKES_IDNetworkMatchArray(intMatchingGBLAKES_IDs)
            ReDim Preserve lonOrderMatchArray(intMatchingGBLAKES_IDs)
        End If
    Next
    intMatchingGBLAKES_IDs = intMatchingGBLAKES_IDs - 1
    ReDim Preserve lonGBLAKES_IDNetworkMatchArray(intMatchingGBLAKES_IDs)
    ReDim Preserve lonOrderMatchArray(intMatchingGBLAKES_IDs)
    Label4.Caption = "Network " & lonChosenNetwork & " has " & intMatchingGBLAKES_IDs + 1 & " catchments."
    DoEvents
    CalcCatchNetRship
    optNetwork.Enabled = True
    If optCatchment Then
        cmdZoomToSelected.Caption = "Zoom to catchment " & lonChosenGBLAKES_ID
        cmdHighlightSelected.Caption = "Highlight catchment " & lonChosenGBLAKES_ID
        cmdCalcTP.Caption = "Calculate total P concentration (for catchments connected to " & lonChosenGBLAKES_ID & ")"
        cmdCalcTP.ControlTipText = "Calculate total P concentration or load (for lochs within network)"
    End If
End If

'create an array containing each of the values of various inputs for the selected catchment
Dim pCursor As ICursor
Dim pRow As IRow
Dim lonRowScroller As Long

If blnScenarioLoaded Then
    intCatchP_PField = pCatchPTable_S.Table.FindField("P")
    If intCatchP_PField = -1 Then
        MsgBox "Cannot find field P in the scenario Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_GBLAKES_IDField = pCatchPTable_S.Table.FindField("GBLAKES_ID")
    If intCatchP_GBLAKES_IDField = -1 Then
        MsgBox "Cannot find field GBLAKES_ID in the scenario Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_LCOVDESCField = pCatchPTable_S.Table.FindField("LCOVDESC")
    If intCatchP_LCOVDESCField = -1 Then
        MsgBox "Cannot find field LCOVDESC in the scenario Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_AreaField = pCatchPTable_S.Table.FindField("Area")
    If intCatchP_AreaField = -1 Then
        MsgBox "Cannot find field Area in the scenario Precursor table", vbCritical
        Exit Sub
    End If
Else
    intCatchP_PField = pCatchPTable.Table.FindField("P")
    If intCatchP_PField = -1 Then
        MsgBox "Cannot find field P in the Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_GBLAKES_IDField = pCatchPTable.Table.FindField("GBLAKES_ID")
    If intCatchP_GBLAKES_IDField = -1 Then
        MsgBox "Cannot find field GBLAKES_ID in the Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_LCOVDESCField = pCatchPTable.Table.FindField("LCOVDESC")
    If intCatchP_LCOVDESCField = -1 Then
        MsgBox "Cannot find field LCOVDESC in the Precursor table", vbCritical
        Exit Sub
    End If
    intCatchP_AreaField = pCatchPTable.Table.FindField("Area")
    If intCatchP_AreaField = -1 Then
        MsgBox "Cannot find field Area in the Precursor table", vbCritical
        Exit Sub
    End If
End If

Set pQueryFilt = New QueryFilter
If blnScenarioLoaded Then
    pQueryFilt.SubFields = "GBLAKES_ID,LCOVDESC,P,Area,ScenarioID"
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lonChosenGBLAKES_ID & " and ScenarioID = " & cboScenarioID
    lontblSelectedCatchCatchPRecords = pCatchPTable_S.Table.RowCount(pQueryFilt)
    Set pCursor = pCatchPTable_S.Table.Search(pQueryFilt, False)
Else
    pQueryFilt.SubFields = "GBLAKES_ID,LCOVDESC,P,AREA,Catch_Net"
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lonChosenGBLAKES_ID
    lontblSelectedCatchCatchPRecords = pCatchPTable.Table.RowCount(pQueryFilt)
    Set pCursor = pCatchPTable.Table.Search(pQueryFilt, False)
End If
If lontblSelectedCatchCatchPRecords = 0 Then
    MsgBox "Nothing selected, please choose a water body.", vbInformation
    Exit Sub
End If
ReDim varSelectedCatchmentCatchP(lontblSelectedCatchCatchPRecords - 1, 2)
dblSumLocalInputs = 0
Set pRow = pCursor.NextRow

'#######################################################################################
'Create an array of the inputs for the selected catchment
'initialise the user modified values with the read in, then always use these fields
'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
'#######################################################################################
If blnScenarioLoaded Then
    ReDim arrayCatchPforChosenGBLAKES_ID(pCatchPTable_S.Table.RowCount(pQueryFilt) - 1, 7)
    While Not pRow Is Nothing
        If pRow.Value(intCatchP_PField) > 0 Or pRow.Value(intCatchP_AreaField) > 0 Then 'some user scenario read in data may be zero, must ignore it
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0) = lonChosenGBLAKES_ID
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) = pRow.Value(intCatchP_LCOVDESCField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) = pRow.Value(intCatchP_PField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) = pRow.Value(intCatchP_AreaField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 4) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) / (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) / 10000)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) 'this will contain modified area
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 6) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 4) 'this will contain modified kg/ha
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) 'this will contain revised P
        End If
        Set pRow = pCursor.NextRow
        lonRowScroller = lonRowScroller + 1
    Wend
    
    'load the varPointSource from pPointSourceTable_S
    'this is a load from the data so erase the pointsource
    Erase varPointSource()
    Dim pCursorPointSources As ICursor
    pQueryFilt.SubFields = "GBLAKES_ID,PointSource,Amount,ScenarioID"
    pQueryFilt.WhereClause = "ScenarioID = " & cboScenarioID
    Set pCursorPointSources = pPointSourceTable_S.Table.Search(pQueryFilt, False)
    'check some point sources are being loaded
    If pPointSourceTable_S.Table.RowCount(pQueryFilt) <> 0 Then
        ReDim varPointSource(pPointSourceTable_S.Table.RowCount(pQueryFilt) - 1, 2)
        Set pRow = pCursorPointSources.NextRow
        i = 0
        While Not pRow Is Nothing
            varPointSource(i, 0) = pRow.Value(intSPS_GBLAKES_ID_field)
            varPointSource(i, 1) = pRow.Value(intSPS_Type_field)
            varPointSource(i, 2) = pRow.Value(intSPS_Amount_field)
            i = i + 1
            Set pRow = pCursorPointSources.NextRow
        Wend
    End If
Else
    ReDim arrayCatchPforChosenGBLAKES_ID(pCatchPTable.Table.RowCount(pQueryFilt) - 1, 7)
    While Not pRow Is Nothing
        If pRow.Value(intCatchP_PField) > 0 Or pRow.Value(intCatchP_AreaField) > 0 Then 'some user scenario read in data may be zero, must ignore it
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 0) = lonChosenGBLAKES_ID
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 1) = pRow.Value(intCatchP_LCOVDESCField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) = pRow.Value(intCatchP_PField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) = pRow.Value(intCatchP_AreaField)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 4) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) / (arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) / 10000)
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 5) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 3) 'this will contain modified area
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 6) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 4) 'this will contain modified kg/ha
            arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 7) = arrayCatchPforChosenGBLAKES_ID(lonRowScroller, 2) 'this will contain revised P
        End If
        Set pRow = pCursor.NextRow
        lonRowScroller = lonRowScroller + 1
    Wend
    
    'load the varPointSource from pPointSourceTable - non-scenario (added 11.08.2015)
    'this is a load from the data so erase the pointsource
    Erase varPointSource()
    'Erase pCursorPointSources()
    Set pCursorPointSources = Nothing
    pQueryFilt.SubFields = "GBLAKES_ID,PointSource,Amount"
    'pQueryFilt.WhereClause = "ScenarioID = " & cboScenarioID
    Set pCursorPointSources = pPointSourceTable.Table.Search(pQueryFilt, False)
    'check some point sources are being loaded
    If pPointSourceTable.Table.RowCount(pQueryFilt) <> 0 Then
        ReDim varPointSource(pPointSourceTable.Table.RowCount(pQueryFilt) - 1, 2)
        Set pRow = pCursorPointSources.NextRow
        i = 0
        While Not pRow Is Nothing
            varPointSource(i, 0) = pRow.Value(intPS_GBLAKES_ID_field)
            varPointSource(i, 1) = pRow.Value(intPS_Type_field)
            varPointSource(i, 2) = pRow.Value(intPS_Amount_field)
            i = i + 1
            Set pRow = pCursorPointSources.NextRow
        Wend
    End If
    
End If
lonRowScroller = 0

'set varModifiedLCoverCoeff() to match the size of arrayCatchPforChosenGBLAKES_ID - this will be used to store the user entered coefficients
ReDim varModifiedLCoverCoeff(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 1)

cmdCalcTP.Enabled = True
lvwCatchmentRelationships2.Height = 0
Frame_Modify_Load.Visible = False
Frame_Modify_Load.Caption = "Modify point sources for " & lonChosenGBLAKES_ID
frameModifyInputs.Visible = False
lvwCatchmentInfo.Visible = False

Frame2UserLandCover.Visible = True

'make the scenario saving option available now that there is a catchment chosen
InitialiseScenarioSave
chkProduceResultsTable.Enabled = False
chkProduceResultsTable.Value = False
chkProduceResultsCSV.Enabled = False
chkProduceResultsCSV.Value = False
cmdZoomToSelected.Enabled = True
cmdHighlightSelected.Enabled = True

DoEvents

'erase - BatchProcessing enhancements
Erase varSplit
'Set vbNullString = Nothing
Set pQueryFilt = Nothing
Set pCursor = Nothing
Set pRow = Nothing
'Set lonRowScroller = Nothing
Set pCursorPointSources = Nothing


End Sub
Private Sub cmdLoad_Data_Click()
Frame_Modify_Load.Visible = False
frameModifyInputs.Visible = False
lvwCatchmentInfo.Visible = False

Dim lonRowsInTable As Long
Dim pRow As IRow
Dim pRowQuery As IQueryFilter
Set pRowQuery = New QueryFilter
Dim pCursor As ICursor
Dim intGBLAKES_IDField As Integer
Dim intSitenameField As Integer
Dim intOrderField As Integer
Dim intNetworkField As Integer
Dim pField As IField
Dim intTempCounter As Integer
Dim strStoredSelectedCatchmentValue As String
Dim blnFound As Boolean

blnDataLoadedFromAScenario = False 'this is used in the creation of a scenario from a scenario
blnScenarioLoaded = False   'this is used in determining output text

'ensure no duplicate data is created
Erase lonGBLAKES_IDNetworkMatchArray()
Erase lonOrderMatchArray()

CreateGBLakesTable
GetPointSourceIndices

'get the selected catchment, SEPA data and flow routing tables from the cbos
'find the catchment layer
FindFLayer pFLayerCatchment, cboCatchment.Text, blnFound
Set pFClassCatchment = pFLayerCatchment.FeatureClass
Set pTableCatchment = pFClassCatchment

'list the fields
Dim pFields As IFields
Set pFields = pTableCatchment.Fields

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter    'get all the records
pQueryFilt.SubFields = "GBLAKES_ID,SiteName,Order_Catch_Net"    'only on these fields to speed up processing
lonNumGBLAKES_IDs = pTableCatchment.RowCount(pQueryFilt)

'get the currently selected catchment and retain it
If Len(cboGBLAKES_IDs.Text) > 0 Then
    strStoredSelectedCatchmentValue = cboGBLAKES_IDs.Text
End If

'the three fields that we are interested in initially are:
'pFieldCatchmentGBLAKES_ID pFieldCatchmentSiteName pFieldCatchmentNetwork
intGBLAKES_IDField = pTableCatchment.FindField("GBLAKES_ID")
If intGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID", vbCritical
End If
intSitenameField = pTableCatchment.FindField("SiteName")
If intSitenameField = -1 Then
    MsgBox "Cannot find field SiteName", vbCritical
End If
intOrderField = pTableCatchment.FindField("Order_")
If intOrderField = -1 Then
    MsgBox "Cannot find field Order_", vbCritical
End If
intNetworkField = pTableCatchment.FindField("Catch_Net")
If intNetworkField = -1 Then
    MsgBox "Cannot find field Catch_Net", vbCritical
End If

Dim lonRowScroller As Long
Dim iOIDList() As Long
Erase iOIDList()
Erase lonGBLAKES_IDArray()
Erase strSitenameArray()
Erase intOrderArray()
Erase lonNetworkArray()
ReDim iOIDList(lonNumGBLAKES_IDs)
ReDim lonGBLAKES_IDArray(lonNumGBLAKES_IDs)
ReDim strSitenameArray(lonNumGBLAKES_IDs)
ReDim intOrderArray(lonNumGBLAKES_IDs)
ReDim lonNetworkArray(lonNumGBLAKES_IDs)

For lonRowScroller = 0 To lonNumGBLAKES_IDs ' - 1
    iOIDList(lonRowScroller) = lonRowScroller
Next

'This error can occur if the database gets to its maximum size!
'compact it and all will be well
Set pCursor = pTableCatchment.GetRows(iOIDList, True)
Set pRow = pCursor.NextRow
Dim lonGBLakesID As Long
lonRowScroller = 0
While Not pRow Is Nothing
'populate the arrays that we will use for subsequent processing
    lonGBLAKES_IDArray(lonRowScroller) = pRow.Value(intGBLAKES_IDField)
    strSitenameArray(lonRowScroller) = pRow.Value(intSitenameField)
    intOrderArray(lonRowScroller) = pRow.Value(intOrderField)
    lonNetworkArray(lonRowScroller) = pRow.Value(intNetworkField)
    Set pRow = pCursor.NextRow
    lonRowScroller = lonRowScroller + 1
Wend

'time - this For loop takes 2 secs to process on a fast PC
Dim i As Long
For i = 0 To lonNumGBLAKES_IDs - 1
    'Note: the use of strSitenameArray(i) rather than ReturnSitename(lonGBLAKES_IDArray(i)) would speed things up
    'significantly, however ReturnSitename() also fixes a problem with Gaelic characters not displaying correctly,
    'partly why it takes so long - it parses the characters individually
    cboGBLAKES_IDs.AddItem lonGBLAKES_IDArray(i) & " - " & ReturnSitename(lonGBLAKES_IDArray(i))  'strSitenameArray(i)
    'add the equivalent WBID to the left drop down menu
    lonGBLakesID = ReturnWFD_WB_ID(CLng(lonGBLAKES_IDArray(i)))
    If lonGBLakesID <> 0 Then   'if no match then it returns a 0
        cboWBID.AddItem lonGBLakesID & " - " & ReturnSitename(lonGBLAKES_IDArray(i))
    End If
Next

Set pCursor = pTableCatchment.GetRows(iOIDList, True) ' back to the beginning
Set pRow = pCursor.NextRow

'get the selected flow table and read into an array
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblFlowRouting Then
        Set pFlowRoutingTable = pTabColl.StandaloneTable(intTempCounter)
    End If
Next

'#######################################################################################
'Process arrayFlowRouting()
'#######################################################################################
Dim intRouteGBLAKES_IDField As Integer
Dim intDSGBLAKES_IDField As Integer
Dim lonRowScroller2 As Long
Dim pQueryFilt2 As IQueryFilter2
Set pQueryFilt2 = New QueryFilter    'don't give it any value yet
lonFlowRoutes = pFlowRoutingTable.Table.RowCount(pQueryFilt2)

Erase arrayFlowRouting()
ReDim arrayFlowRouting(lonFlowRoutes, 1)
Dim lonFlowRowScroller2 As Long
Dim iOIDFlowList() As Long
ReDim iOIDFlowList(lonFlowRoutes)

For lonRowScroller2 = 1 To lonFlowRoutes    'this should be 1 - the first record in the table is 1, the last is 4485
    iOIDFlowList(lonRowScroller2) = lonRowScroller2
Next

intRouteGBLAKES_IDField = pFlowRoutingTable.Table.FindField("GBLAKES_ID")
If intRouteGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID", vbCritical
    Exit Sub
End If
intDSGBLAKES_IDField = pFlowRoutingTable.Table.FindField("DownstreamGBLAKES_ID")
If intDSGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field DS_SITECOD", vbCritical
    Exit Sub
End If

Set pCursor = Nothing
Set pCursor = pFlowRoutingTable.Table.GetRows(iOIDFlowList, True)

If Not pFlowRoutingTable.Table.HasOID Then
    MsgBox pFlowRoutingTable.Name & " is not registered. In ArcCatalog right click the table and choose Register with GDB, do the same with all tables.", vbCritical
    Exit Sub
End If

Set pRow = pCursor.NextRow
lonRowScroller2 = 0
While Not pRow Is Nothing
'populate the arrays that we will use for subsequent processing
    arrayFlowRouting(lonRowScroller2, 0) = pRow.Value(intRouteGBLAKES_IDField)
    arrayFlowRouting(lonRowScroller2, 1) = pRow.Value(intDSGBLAKES_IDField)
    Set pRow = pCursor.NextRow
    lonRowScroller2 = lonRowScroller2 + 1
Wend
    
'populate the combobox with the shapefiles that the user may want to choose to overlay
'initialise the catchment
Dim pCompositeLayer As ICompositeLayer
Dim j As Integer

For intTempCounter = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(intTempCounter) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(intTempCounter)
        For j = 0 To pCompositeLayer.Count - 1
        If TypeOf pCompositeLayer.Layer(j) Is IFeatureLayer Then
            If pCompositeLayer.Layer(j).Name <> cboCatchment.Value And pCompositeLayer.Layer(j).Name <> strTblFlowRouting Then
                cboUserShapefile.AddItem pCompositeLayer.Layer(j).Name
                cboUserShapefile.Text = pCompositeLayer.Layer(j).Name
                cboSlopeLCoverShapefile.AddItem pCompositeLayer.Layer(j).Name
                If pCompositeLayer.Layer(j).Name = strShapefileSlopeClass_LandCover Then
                    cboSlopeLCoverShapefile.Value = pCompositeLayer.Layer(j).Name   'for intersecting with user polygons
                End If
            End If
        End If
        Next j
    Else ' not grouped layer
        If TypeOf pMap.Layer(intTempCounter) Is IFeatureLayer Then
            If DBExists(pMap.Layer(intTempCounter)) Then
                If pMap.Layer(intTempCounter).Name <> cboCatchment.Value And pMap.Layer(intTempCounter).Name <> strTblFlowRouting Then
                    cboUserShapefile.AddItem pMap.Layer(intTempCounter).Name
                    If cboUserShapefile.Text = "" Then
                            cboUserShapefile.Text = pMap.Layer(intTempCounter).Name
                        Else
                            If pMap.Layer(intTempCounter).Name Like "*User*" Then
                                cboUserShapefile.Text = pMap.Layer(intTempCounter).Name
                            End If
                    End If
                    
                    cboSlopeLCoverShapefile.AddItem pMap.Layer(intTempCounter).Name
                    If pMap.Layer(intTempCounter).Name = strShapefileSlopeClass_LandCover Then
                        cboSlopeLCoverShapefile.Value = pMap.Layer(intTempCounter).Name
                    End If
                End If
            End If
            If ItExists(pMap.Layer(intTempCounter)) Then
                If pMap.Layer(intTempCounter).Name <> cboCatchment.Value And pMap.Layer(intTempCounter).Name <> strTblFlowRouting Then
                    cboUserShapefile.AddItem pMap.Layer(intTempCounter).Name
                    cboUserShapefile.Text = pMap.Layer(intTempCounter).Name
                    cboSlopeLCoverShapefile.AddItem pMap.Layer(intTempCounter).Name
                    If pMap.Layer(intTempCounter).Name = strShapefileSlopeClass_LandCover Then
                        cboSlopeLCoverShapefile.Value = pMap.Layer(intTempCounter).Name
                    End If
                End If
            End If
        End If
    End If
Next

cmdGetCatchmentInfo.Enabled = True
cmdIntersectUserShapefile.Enabled = True
cboGBLAKES_IDs.Enabled = True
cboWBID.Enabled = True
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label49.Enabled = True
Label50.Enabled = True
Label51.Enabled = True
cboUserShapefile.Enabled = True
cboSlopeLCoverShapefile.Enabled = True
txtOutputShapefileName.Enabled = True
optCatchment.Enabled = True
lblScenarioSaveWarning.Visible = True
lblScenarioSaveWarning2.Visible = True
lblReportSaveWarning.Visible = True
Frame1Scenario.Visible = False
Frame2UserLandCover.Visible = False
chkProduceResultsTable.Enabled = False
chkProduceResultsTable.Value = False
chkProduceResultsCSV.Enabled = False
chkProduceResultsCSV.Value = False

'get the SEPA classification data and read into an array
LoadSepaMonitoringIntoArray

'get the SEPA classification concentration statistic data and read into an array
LoadSepaClassConcStat

SwitchReportToBaseline
blnNoLongerBaseline = False

cboWBID_Click

End Sub
Private Sub cmdHighlightSelected_Click()
'#######################################################################################
'Highlight the catchment selected in cboGBLAKES_IDs using the layer in cboCatchment - don't use earlier found layer in case the selection has been redone
'get the selected catchment and flow routing tables from the cbos
'find the catchment layer
'#######################################################################################

'Switch to data view
If pMxDoc.ActiveView Is pMxDoc.PageLayout Then
    Set pMxDoc.ActiveView = pMxDoc.Maps.Item(0)
End If

FindFLayer pFLayerCatchment, cboCatchment.Text, True
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView
Dim varSplit As Variant

If cboGBLAKES_IDs.Text <> "" Then
    varSplit = Split(cboGBLAKES_IDs.Text, " - ")
    lonChosenGBLAKES_ID = varSplit(0)
    strChosenSitename = varSplit(1)
End If

Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment
If optCatchment Then
    pQueryFilter.WhereClause = "GBLAKES_ID =" & lonChosenGBLAKES_ID
    Else
'zoom to network lonChosenNetwork
    pQueryFilter.WhereClause = "NETWORK =" & lonChosenNetwork
End If

'Display the selection
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
pFeatureselection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing

End Sub
Sub DisplaySewageInfo()
'#######################################################################################
'Get data from PerCapitaTPLoads and CatchmentSewage
'#######################################################################################
ReDim arrayPerCapitaTPLoads(1, 1)

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
Dim pCursor As ICursor
Dim pRow As IRow

If blnScenarioLoaded Then
    pQueryFilt.WhereClause = "ScenarioID = " & lonSelectedScenario
    Set pCursor = pPerCapitaTPLoads_S.Table.Search(pQueryFilt, False)
    intPerCapitaTPLoadField = pPerCapitaTPLoads_S.Table.FindField("PerCapitaTPLoad")
    intPerCapitaUrb_RurField = pPerCapitaTPLoads_S.Table.FindField("Urb_Rur")
Else
    Set pCursor = pPerCapitaTPLoads.Table.Search(pQueryFilt, False)
End If

Set pRow = pCursor.NextRow
Dim i As Integer
i = 0

While Not pRow Is Nothing
    arrayPerCapitaTPLoads(i, 0) = pRow.Value(intPerCapitaUrb_RurField)
    arrayPerCapitaTPLoads(i, 1) = pRow.Value(intPerCapitaTPLoadField)
    i = 1
    Set pRow = pCursor.NextRow
Wend
For i = 0 To 1
    If arrayPerCapitaTPLoads(i, 0) = "Urban" Then
        If arrayPerCapitaTPLoads(i, 1) < 1 Then
            lblUrbanLoad.Caption = "Urban = 0" & Format(arrayPerCapitaTPLoads(i, 1), "#.000")
            dblUrbanPerCapitaTPLoad = arrayPerCapitaTPLoads(i, 1)
        Else
            lblUrbanLoad.Caption = "Urban = " & Format(arrayPerCapitaTPLoads(i, 1), "#.000")
            dblUrbanPerCapitaTPLoad = arrayPerCapitaTPLoads(i, 1)
        End If
    ElseIf arrayPerCapitaTPLoads(i, 0) = "Rural" Then
        If arrayPerCapitaTPLoads(i, 1) < 1 Then
            lblRuralLoad.Caption = "Rural = 0" & Format(arrayPerCapitaTPLoads(i, 1), "#.000")
            dblRuralPerCapitaTPLoad = arrayPerCapitaTPLoads(i, 1)
        Else
            lblRuralLoad.Caption = "Rural = " & Format(arrayPerCapitaTPLoads(i, 1), "#.000")
            dblRuralPerCapitaTPLoad = arrayPerCapitaTPLoads(i, 1)
        End If
    End If
Next

lblUrbanPop.Caption = "Urban = " & CatchNetRship(intIndexSelectedGBLAKES_ID, 18)
lblRuralPop.Caption = "Rural = " & CatchNetRship(intIndexSelectedGBLAKES_ID, 19)

End Sub
Sub GetFieldIndices()

intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff")
'intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff_M10")
'intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff_M5")
'intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff_P5")
'intLocalRunoffField = pLoadPrecursorTable.Table.FindField("LocalRunoff_P10")
If intLocalRunoffField = -1 Then
    MsgBox "Cannot find field LocalRunoff", vbCritical
    Exit Sub
End If
intLocalAreaField = pLoadPrecursorTable.Table.FindField("LocalArea")
If intLocalAreaField = -1 Then
    MsgBox "Cannot find field LocalArea", vbCritical
    Exit Sub
End If
intPrecursorGBLAKES_IDField = pLoadPrecursorTable.Table.FindField("GBLAKES_ID")
If intPrecursorGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the Precursor table", vbCritical
    Exit Sub
End If
intLochDepthField = pLoadPrecursorTable.Table.FindField("LochMeanDepth")
If intLochDepthField = -1 Then
    MsgBox "Cannot find field LochMeanDepth in the Precursor table", vbCritical
    Exit Sub
End If
intLochAreaField = pLoadPrecursorTable.Table.FindField("LochArea")
If intLochAreaField = -1 Then
    MsgBox "Cannot find field LochArea in the Precursor table", vbCritical
    Exit Sub
End If
intOECDDenominatorField = pLoadPrecursorTable.Table.FindField("OECDDenominator")
If intOECDDenominatorField = -1 Then
    MsgBox "Cannot find field OECDDenominator in the Precursor table", vbCritical
    Exit Sub
End If
intOECDExponentDenominatorField = pLoadPrecursorTable.Table.FindField("OECDExponentDenominator")
If intOECDExponentDenominatorField = -1 Then
    MsgBox "Cannot find field OECDExponentDenominator in the Precursor table", vbCritical
    Exit Sub
End If
intOrderField = pLoadPrecursorTable.Table.FindField("LochOrder")
If intOrderField = -1 Then
    MsgBox "Cannot find field LochOrder in the Precursor table", vbCritical
    Exit Sub
End If
intSewageGBLAKES_IDField = pCatchmentSewageTable.Table.FindField("GBLAKES_ID")
If intSewageGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the " & strTblCatchmentSewageName & " table", vbCritical
    Exit Sub
End If
intUrb_RurField = pCatchmentSewageTable.Table.FindField("Urb_Rur")
If intUrb_RurField = -1 Then
    MsgBox "Cannot find field Urb_Rur in the " & strTblCatchmentSewageName & " table", vbCritical
    Exit Sub
End If
intLoadField = pCatchmentSewageTable.Table.FindField("Load")
'intLoadField = pCatchmentSewageTable.Table.FindField("NoUrb_Load")
If intLoadField = -1 Then
    MsgBox "Cannot find field Load in the " & strTblCatchmentSewageName & " table", vbCritical
    Exit Sub
End If
intPopulationField = pCatchmentSewageTable.Table.FindField("Pop")
'intPopulationField = pCatchmentSewageTable.Table.FindField("NoUrb_Pop")
If intPopulationField = -1 Then
    MsgBox "Cannot find field Pop in the " & strTblCatchmentSewageName & " table", vbCritical
    Exit Sub
End If
intCatchP_PField = pCatchPTable.Table.FindField("P")
If intCatchP_PField = -1 Then
    MsgBox "Cannot find field P in the Precursor table", vbCritical
    Exit Sub
End If
intCatchP_GBLAKES_IDField = pCatchPTable.Table.FindField("GBLAKES_ID")
If intCatchP_GBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the Precursor table", vbCritical
    Exit Sub
End If
intCatchP_LCOVDESCField = pCatchPTable.Table.FindField("LCOVDESC")
If intCatchP_LCOVDESCField = -1 Then
    MsgBox "Cannot find field LCOVDESC in the tblCatchP table", vbCritical
    Exit Sub
End If
intCatchP_AreaField = pCatchPTable.Table.FindField("Area")
If intCatchP_AreaField = -1 Then
    MsgBox "Cannot find field Area in the tblCatchP table", vbCritical
    Exit Sub
End If
intPerCapitaUrb_RurField = pPerCapitaTPLoads.Table.FindField("Urb_Rur")
If intPerCapitaUrb_RurField = -1 Then
    MsgBox "Cannot find field Urb_Rur in " & strTblPerCapitaTPLoadsName, vbCritical
    Exit Sub
End If
intPerCapitaTPLoadField = pPerCapitaTPLoads.Table.FindField("PerCapitaTPLoad")
If intPerCapitaTPLoadField = -1 Then
    MsgBox "Cannot find field PerCapitaTPLoad in " & strTblPerCapitaTPLoadsName, vbCritical
    Exit Sub
End If
intLowerDensityField = pPerCapitaTPLoads.Table.FindField("LowerDensity")
If intLowerDensityField = -1 Then
    MsgBox "Cannot find field LowerDensity in " & strTblPerCapitaTPLoadsName, vbCritical
    Exit Sub
End If
intUpperDensityField = pPerCapitaTPLoads.Table.FindField("UpperDensity")
If intUpperDensityField = -1 Then
    MsgBox "Cannot find field UpperDensity in " & strTblPerCapitaTPLoadsName, vbCritical
    Exit Sub
End If
intGBLAKES_IDFieldBreakPoints = pTPBreakPoints.Table.FindField("GBLAKES_ID")
If intGBLAKES_IDFieldBreakPoints = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
intReference_TypeField = pTPBreakPoints.Table.FindField("Reference_Type")
If intReference_TypeField = -1 Then
    MsgBox "Cannot find field Reference_Type in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
intHighGood_PField = pTPBreakPoints.Table.FindField("HighGood_P")
If intHighGood_PField = -1 Then
    MsgBox "Cannot find field HighGood_P in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
intGoodModerate_PField = pTPBreakPoints.Table.FindField("GoodModerate_P")
If intGoodModerate_PField = -1 Then
    MsgBox "Cannot find field GoodModerate_P in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
intModeratePoor_PField = pTPBreakPoints.Table.FindField("ModeratePoor_P")
If intModeratePoor_PField = -1 Then
    MsgBox "Cannot find field ModeratePoor_P in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
intPoorBad_PField = pTPBreakPoints.Table.FindField("PoorBad_P")
If intPoorBad_PField = -1 Then
    MsgBox "Cannot find field PoorBad_P in the " & strTblTPBreakPointsName & " table", vbCritical
    Exit Sub
End If
End Sub
Sub GetPointSourceIndices()
'get the non-scenario point source indices
intPS_GBLAKES_ID_field = pPointSourceTable.Table.FindField("GBLAKES_ID")
If intPS_GBLAKES_ID_field = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
intPS_Type_field = pPointSourceTable.Table.FindField("PointSource")
If intPS_Type_field = -1 Then
    MsgBox "Cannot find field PointSource in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
intPS_Amount_field = pPointSourceTable.Table.FindField("Amount")
If intPS_Amount_field = -1 Then
    MsgBox "Cannot find field Amount in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
End Sub
Sub GetScenarioPointSourceIndices()
'get the scenario point source indices
intSPS_GBLAKES_ID_field = pPointSourceTable_S.Table.FindField("GBLAKES_ID")
If intSPS_GBLAKES_ID_field = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
intSPS_Type_field = pPointSourceTable_S.Table.FindField("PointSource")
If intSPS_Type_field = -1 Then
    MsgBox "Cannot find field PointSource in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
intSPS_Amount_field = pPointSourceTable_S.Table.FindField("Amount")
If intSPS_Amount_field = -1 Then
    MsgBox "Cannot find field Amount in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
intSPS_ScenarioID_field = pPointSourceTable_S.Table.FindField("ScenarioID")
If intSPS_ScenarioID_field = -1 Then
    MsgBox "Cannot find field ScenarioID in the " & strScenarioTblPointSource & " table", vbCritical
    Exit Sub
End If
End Sub

Private Sub cmdInfoChangeP_Click()
MsgBox "This option will change the P coefficient for the selected land cover " _
& vbCrLf & "(in this calculation, not the source data) to a single value, regardless " & vbCrLf & " of variations in slope.", vbInformation, "Information"
End Sub

Private Sub cmdIntersectUserShapefile_Click()
'#######################################################################################
'This routine imports a user created shapefile with land cover/slope that differs to the base line data SlopeClass_LandCover
'The table CatchP has been derived from the base line SlopeClass_LandCover so this routine creates scenario data in
'CatchP_S, in addition to a new shapefile containing an updated base line data for the selected network of catchments
'This tool does not update the base line data, it is only used as an input for this routine.
'1. Validate the user entered polygon data - are the attributes/codes correct?
'2. Get the catchment network polygons ### output = pFLayerCatchment.Name & "_Selected"
'3. Use this to clip the base line land cover to a network extent base line land cover ### output = C:\temp\CatchClipSlpClsLandCover.shp
'4. Overlap/update the base line polygons using the user data ### output = txtOutputShapefileName.Text & ".shp"
'5. Recalculate the areas in the new user modified network land cover
'6. Update the CatchP table (This was "Create a scenario SlopeClass_LandCover table" prior to Jully '11)
'Note that this tool must use ESRI Geoprocessing and that this is known to crash. Please exit ArcGIS and try again if repeat
'crashes occur
'
'Now want to use the existing base line data to determine both the AverageExp value and to supply the slope code
'so to determine whether blanket bog & peatland is 103, 203 or 303. At present this data is not accessed - it just
'goes into CatchClipSlpClsLandCover.shp which is then 'Updated' using the user supplied data. As an intermediate step
'I could do a spatial intersect and use the first digit in the codes in the intersect and a LUT to get AverageExp values
'for this new, modified user supplied land cover. Do a table update for AverageExp and then use this to update the baseline
'data.
'#######################################################################################

'get the chosen feature layer from cboUserShapefile
'to ensure everything is calculated for the scenario run CalculateTP

'check that c:\temp exists - used for storing temp data
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\temp") Then
    MsgBox "Warning, please create a folder c:\temp and restart this process.", vbCritical
    Exit Sub
End If

CalculateTP

'attempt to deal with Geoprocessing crashes
Dim pGpUtils As IGPUtilities2
Set pGpUtils = New GPUtilities
'pGpUtils.RemoveInternalData
pGpUtils.ClearInMemoryWorkspace
pGpUtils.ReleaseInternals

'#######################################################################################
'1. Validate the user entered polygon data - does it exist, are the attributes/codes correct?
'#######################################################################################
If txtOutputShapefileName.Value = "" Or IsNull(txtOutputShapefileName) Then
    MsgBox "Please enter a suitable name for your output shapefile.", vbCritical
    Exit Sub
End If

'if the user has added the ".shp" extension delete it, the code doesn't expect it
Dim varSplit As Variant
If txtOutputShapefileName.Text <> "" Then
    varSplit = Split(txtOutputShapefileName.Text, ".")
    If UBound(varSplit) > 0 Then
        If varSplit(1) = "shp" Then
            txtOutputShapefileName = varSplit(0)
        End If
    End If
End If

'check they exist
Dim blnFound1 As Boolean
Dim blnFound2 As Boolean
FindFLayer pFLayerUserLandCoverInput, cboUserShapefile.Value, blnFound1
FindFLayer pFLayerSlopeLandClass, cboSlopeLCoverShapefile.Value, blnFound2
If blnFound1 And blnFound2 Then
        'MsgBox "Good!"
    Else
        MsgBox cboUserShapefile.Value & " and " & cboSlopeLCoverShapefile.Value & " must both be loaded into your Table of Contents to proceed.", vbCritical
        Exit Sub
End If

'If the LocalCatchment_and_Network layer has anything joined to it then the select by command will not work properly
'Prompt the user and then remove the joins
Dim pCatchments As ILayer2
Set pCatchments = pFLayerCatchment

Dim pDispRC As IDisplayRelationshipClass
Set pDispRC = pCatchments
  
Dim strResponse As String
'Remove all joins
If Not pDispRC.RelationshipClass Is Nothing Then
    strResponse = MsgBox(pCatchments.Name & " has one or more layers or tables joined to it. This procedure requires all joins to be removed." & vbCrLf & _
    "Do you want them removed?", vbYesNo)
    If strResponse = vbNo Then
        Exit Sub
    Else
        pDispRC.DisplayRelationshipClass Nothing, esriLeftInnerJoin
        MsgBox "Joins removed."
    End If
End If

'check the cboCatchmentScenario.text is loaded
Dim pFLayerTemp As IFeatureLayer
FindFLayer pFLayerTemp, cboCatchmentScenario.Value, blnFound1
If Not blnFound1 Then
    MsgBox "You must have a shapefile selected in the scenario tab to save your scenario. '" & vbCrLf & strScenarioLocalCatchmentAndNetwork & "' is the expected name.", vbCritical
    Exit Sub
End If

'Check the attributes of the user polygon file, it must have a field GridCode and a field AverageExp. These are defined in 'Activate'
Dim intFieldGRIDCODE As Integer
Dim intFieldAverageExport As Integer
Dim pFClassUserLandCoverInput As IFeatureClass
Dim pTableUserLandCoverInput As ITable
Set pFClassUserLandCoverInput = pFLayerUserLandCoverInput.FeatureClass
Set pTableUserLandCoverInput = pFClassUserLandCoverInput
intFieldGRIDCODE = pTableUserLandCoverInput.FindField(strFieldGRIDCODE)
If intFieldGRIDCODE = -1 Then
    MsgBox "Cannot find field " & strFieldGRIDCODE, vbCritical
    Exit Sub
End If

'ensure that GridCode is a long and AverageExp a double
If pTableUserLandCoverInput.Fields.Field(intFieldGRIDCODE).Type <> esriFieldTypeInteger Then
    MsgBox "The field " & strFieldGRIDCODE & " must be a long integer.", vbCritical
    Exit Sub
End If

'#######################################################################################
'2. Get the catchment network polygons (use the select method already developed)
'#######################################################################################
'find the catchment layer
FindFLayer pFLayerCatchment, cboCatchment.Text, True
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView
Dim varSplit2 As Variant

If cboGBLAKES_IDs.Text <> "" Then
    varSplit2 = Split(cboGBLAKES_IDs.Text, " - ")
    lonChosenGBLAKES_ID = varSplit2(0)
    strChosenSitename = varSplit2(1)
End If

Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment
pQueryFilter.WhereClause = "NETWORK =" & lonChosenNetwork
pFeatureselection.SelectFeatures pQueryFilter, esriSelectionResultNew, False

Dim pSelectedCatchmentsLayer As ILayer
Dim pSelectedCatchmentsLayerDefinition As IFeatureLayerDefinition
Set pSelectedCatchmentsLayerDefinition = pFLayerCatchment

Set pSelectedCatchmentsLayer = pSelectedCatchmentsLayerDefinition.CreateSelectionLayer(pFLayerCatchment.Name & "_Selected", True, "", "")
pMap.AddLayer pSelectedCatchmentsLayer

'#######################################################################################
'3. Use this selection to clip the base line land cover to a network extent base line land cover
'#######################################################################################
Label37.Caption = "Processing... Please wait for the 'Done!' message. This will take a few minutes (the base line slope class and land cover data set is very large)."
DoEvents

Dim pParams As IVariantArray
Set pParams = New varArray
'Input feature layer name in TOC
pParams.Add pFLayerSlopeLandClass.FeatureClass.AliasName
'Clip feature layer name in TOC
pParams.Add pSelectedCatchmentsLayer.Name
'Output feature class name and location
pParams.Add "C:\temp\CatchClipSlpClsLandCover.shp"  'this is an intermediate file that is deleted at the end of this command

'Execute the tool
Dim pGP As IGeoProcessor
Set pGP = New GeoProcessor
pGP.LogHistory = True

Dim pGPResults As IGeoProcessorResult
Set pGPResults = Nothing
'On Error GoTo errorhandler
'to do - put an "On Error Resume Next" here to handle the error better
'#######################################################################################################
'#######################################################################################
Set pGPResults = pGP.Execute("Clip_analysis", pParams, Nothing) 'this will crash at times. It's a known bug in ArcGIS, exit and re-run
'#######################################################################################
'#######################################################################################################
Label37.Caption = "Existing slope class/land use polygons clipped to temporary file..."
DoEvents

'#######################################################################################
'3a. Intersect pFLayerUserLandCoverInput and CatchClipSlpClsLandCover to get access to
'the underlying slopes and therefore AverageExp of source data
'#######################################################################################
Set pParams = New varArray
'add input feature
pParams.Add cboUserShapefile.Value & ";" & "C:\temp\CatchClipSlpClsLandCover.shp"
pParams.Add "C:\temp\CatchClipSlpClsLandCover_User_Inter.shp"
pParams.Add "NO_FID"
pParams.Add "0"
pParams.Add "INPUT"
Dim pGP2 As IGeoProcessor
Set pGP2 = New GeoProcessor
pGP.LogHistory = True

Dim pGPResults2 As IGeoProcessorResult
Set pGPResults2 = Nothing
'#######################################################################################################
Set pGPResults2 = pGP2.Execute("Intersect_analysis", pParams, Nothing) 'this will crash at times, for no reason.
'#######################################################################################################

'#######################################################################################
'3b. Get the leading character from CatchClipSlpClsLandCover_User_Inter.GRIDCODE_1
'and replace the leading character in CatchClipSlpClsLandCover_User_Inter.GridCode
'then use the resulting value to find the correct AverageExp from Exports
'#######################################################################################

Dim pFLayerUserSlope_Inter As IFeatureLayer
Dim pFClassUserSlope_Inter As IFeatureClass
Dim pTableUserSlope_Inter As ITable
FindFLayer pFLayerUserSlope_Inter, "CatchClipSlpClsLandCover_User_Inter", False
Set pFClassUserSlope_Inter = pFLayerUserSlope_Inter.FeatureClass
Set pTableUserSlope_Inter = pFClassUserSlope_Inter
Dim strTemp As String
Dim pFeature As IFeature
Dim pFCursor As IFeatureCursor
Set pFCursor = pTableUserSlope_Inter.Update(Nothing, True)
Set pFeature = pFCursor.NextFeature
While Not pFeature Is Nothing
    pFeature.Value(pFClassUserSlope_Inter.FindField("GridCode")) = Left(pFeature.Value(pFClassUserSlope_Inter.FindField("GRIDCODE_1")), 1) & Right(pFeature.Value(pFClassUserSlope_Inter.FindField("GridCode")), 2)
    pFeature.Value(pFClassUserSlope_Inter.FindField("AverageExp")) = ReturnExportCoeff(varArrayExportsTable, CInt(Left(pFeature.Value(pFClassUserSlope_Inter.FindField("GRIDCODE_1")), 1) & Right(pFeature.Value(pFClassUserSlope_Inter.FindField("GridCode")), 2)))
    pFCursor.UpdateFeature pFeature
    Set pFeature = pFCursor.NextFeature
Wend

Label37.Caption = "Slope class/land use coefficients obtained from exports source data..."
DoEvents

'#######################################################################################
'4. Overlap/update the base line polygons using the user data (which has been modified using baseline slope data
'#######################################################################################
Set pParams = New varArray
'add input feature
pParams.Add "C:\temp\CatchClipSlpClsLandCover.shp"
pParams.Add "C:\temp\CatchClipSlpClsLandCover_User_Inter.shp"
pParams.Add "c:\temp\updated_baseline_landcover.shp"   'this is the intermediate out shapefile - but no GBLAKES_ID - that comes in (6)
'add borders boolean
pParams.Add "True"
'add cluster tol
pParams.Add "0.0"

Dim pGP3 As IGeoProcessor
Dim pGPResults3 As IGeoProcessorResult
Set pGP3 = New GeoProcessor
Set pGPResults3 = Nothing
'#######################################################################################################
'#######################################################################################
Set pGPResults3 = pGP3.Execute("Update_analysis", pParams, Nothing) 'this will crash at times - ArcGIS bug.
'#######################################################################################
'#######################################################################################################

Set pGP = Nothing
Set pGPResults = Nothing
Set pGP2 = Nothing
Set pGPResults2 = Nothing
Set pGP3 = Nothing
Set pGPResults3 = Nothing
Set pParams = Nothing

Label37.Caption = "Produced updated baseline slope/land cover..."
DoEvents

'#######################################################################################
'5. Recalculate the areas in the new user modified network land cover
'#######################################################################################
'get the table for txtOutputShapefileName.Text
Dim pFModifiedSlopeCoverLayer As IFeatureLayer2
Dim blnFoundSlopeCover As Boolean
FindFLayer pFModifiedSlopeCoverLayer, "updated_baseline_landcover", blnFoundSlopeCover
If Not blnFoundSlopeCover Then
    MsgBox "updated_baseline_landcover.shp not found.", vbCritical
    Label37.Caption = "Error encountered - updated_baseline_landcover.shp not found."
    DoEvents
    Exit Sub
End If

'get the table for pFModifiedSlopeCoverLayer and update the area field
Dim pFCModifiedSlopeCoverLayer As IFeatureClass
Set pFCModifiedSlopeCoverLayer = pFModifiedSlopeCoverLayer.FeatureClass

'there is an ArcGIS bug that causes modeless forms to exit when Calculate is used, I want to use modeless so this method...
Set pFCursor = pFCModifiedSlopeCoverLayer.Update(Nothing, True)
Set pFeature = pFCursor.NextFeature
Dim pPolygon As IPolygon

Label37.Caption = "Summarised the areas in slope/land cover... Creating the scenario data"
DoEvents

'#######################################################################################
'6. Create a scenario SlopeClass_LandCover table from txtOutputShapefileName.Text
'#######################################################################################
'the output goes into CatchP_S - the scenario version
'this will be what the user must load when he wants to implement his land cover in scenarios

Set pParams = Nothing
Set pParams = New varArray

'add input features
pParams.Add pFLayerCatchment.FeatureClass.AliasName & ";" & pFModifiedSlopeCoverLayer.FeatureClass.AliasName
'add output feature
pParams.Add "c:\temp\" & txtOutputShapefileName.Text & ".shp"  'this output has GBLAKES_ID, Gridcode, AverageExp - need to recalc area
Set pGP3 = New GeoProcessor
Set pGPResults3 = Nothing
Set pGPResults3 = pGP3.Execute("Intersect_analysis", pParams, Nothing)

Set pGPResults3 = Nothing
Set pGP3 = Nothing
Set pParams = Nothing
Set pFCursor = Nothing

'want only GBLAKES_ID, LCOVDESC, P, Area, Catch_Net
Dim pFLayerNewCatchSlopeLCover_Inter As IFeatureLayer2
FindFLayer pFLayerNewCatchSlopeLCover_Inter, txtOutputShapefileName.Text, True
Dim pFClassNewCatchSlopeLCover_Inter As IFeatureClass
Set pFClassNewCatchSlopeLCover_Inter = pFLayerNewCatchSlopeLCover_Inter.FeatureClass
DeleteField pFClassNewCatchSlopeLCover_Inter, "SiteName"
DeleteField pFClassNewCatchSlopeLCover_Inter, "Order_"
DeleteField pFClassNewCatchSlopeLCover_Inter, "HistoricNu"
DeleteField pFClassNewCatchSlopeLCover_Inter, "NutrientCl"
DeleteField pFClassNewCatchSlopeLCover_Inter, "HistoricMo"
DeleteField pFClassNewCatchSlopeLCover_Inter, "PercentSew"
DeleteField pFClassNewCatchSlopeLCover_Inter, "FID_update"
DeleteField pFClassNewCatchSlopeLCover_Inter, "ID"
DeleteField pFClassNewCatchSlopeLCover_Inter, "FID_LocalC"
DeleteField pFClassNewCatchSlopeLCover_Inter, "LochArea"
DeleteField pFClassNewCatchSlopeLCover_Inter, "LochMeanDe"
DeleteField pFClassNewCatchSlopeLCover_Inter, "LochMean_1"
DeleteField pFClassNewCatchSlopeLCover_Inter, "TPRatio"
DeleteField pFClassNewCatchSlopeLCover_Inter, "CurrentMod"
DeleteField pFClassNewCatchSlopeLCover_Inter, "NETWORK"
DeleteField pFClassNewCatchSlopeLCover_Inter, "Shape_Leng"
DeleteField pFClassNewCatchSlopeLCover_Inter, "Shape_Area"
DeleteField pFClassNewCatchSlopeLCover_Inter, "Shape_Le_1"
DeleteField pFClassNewCatchSlopeLCover_Inter, "Shape_Ar_1"
 
'calc the area in txtOutputShapefileName.Text
Dim pFCursor2 As IFeatureCursor
Set pFCursor2 = pFClassNewCatchSlopeLCover_Inter.Update(Nothing, True)
Dim pFeature2 As IFeature
Set pFeature2 = pFCursor2.NextFeature
Dim pPolygon2 As IPolygon
Dim dblArea As Double
Dim pArea As IArea

While Not pFeature2 Is Nothing
    Set pPolygon2 = pFeature2.Shape
    Set pArea = pPolygon2
    pFeature2.Value(pFClassNewCatchSlopeLCover_Inter.FindField("Area")) = pArea.Area
    pFCursor2.UpdateFeature pFeature2
    Set pFeature2 = pFCursor2.NextFeature
Wend

'add a LCOVDESC code to txtOutputShapefileName.Text and populate it using Exports
'use strTblExportsName
Dim pQueryFilt As IQueryFilter
Set pQueryFilt = New QueryFilter
pQueryFilt.SubFields = "LCOVCODE,SlopeCode,MatchCode,LCOVDESC,Min_,Max_,Average"
Dim pCursor As ICursor
Set pCursor = pExportsTable.Table.Search(pQueryFilt, False)
Dim pRow As IRow
Set pRow = pCursor.NextRow
Dim lonCounter As Long
lonCounter = 0
While Not pRow Is Nothing
    varArrayExports(lonCounter, 0) = pRow.Value(intMatchCode)
    varArrayExports(lonCounter, 1) = pRow.Value(intLCOVDESC)
    Set pRow = pCursor.NextRow
    lonCounter = lonCounter + 1
Wend

Dim pField As IField
Dim pFieldEdit As IFieldEdit2
Set pFieldEdit = New Field
pFieldEdit.Name = "LCOVDESC"
pFieldEdit.Length = 64
pFieldEdit.Type = 4

Set pField = pFieldEdit
pFClassNewCatchSlopeLCover_Inter.AddField pFieldEdit

Set pFieldEdit = New Field
pFieldEdit.Name = "P"
pFieldEdit.Type = 3

Set pField = pFieldEdit
pFClassNewCatchSlopeLCover_Inter.AddField pFieldEdit

Label37.Caption = "Calculating the P..."
DoEvents

'#######################################################################################
'Calculate the P
'#######################################################################################
Set pFCursor = Nothing
Set pFeature = Nothing
Set pFCursor = pFClassNewCatchSlopeLCover_Inter.Update(Nothing, True)
Set pFeature = pFCursor.NextFeature

While Not pFeature Is Nothing
'area is in metres, units go to hectares
    pFeature.Value(pFClassNewCatchSlopeLCover_Inter.FindField("P")) = pFeature.Value(pFClassNewCatchSlopeLCover_Inter.FindField("Area")) * _
    pFeature.Value(pFClassNewCatchSlopeLCover_Inter.FindField("AverageExp")) / 10000
    pFeature.Value(pFClassNewCatchSlopeLCover_Inter.FindField("LCOVDESC")) = _
    GetLCOVDESC(pFeature.Value(pFClassNewCatchSlopeLCover_Inter.FindField("GRIDCODE")), varArrayExports())
    pFCursor.UpdateFeature pFeature
    Set pFeature = pFCursor.NextFeature
Wend

'note that CatchP has one line per land cover per site, not one line per slope class/land cover class per site.
'so summarise on LCOVDESC
Set pParams = Nothing
Set pParams = New varArray

'add input features
pParams.Add pFClassNewCatchSlopeLCover_Inter.AliasName
pParams.Add "c:\temp\summary_table.dbf"
'add statistics fields
pParams.Add "P SUM;Area SUM"
'add case field
pParams.Add "GBLAKES_ID;LCOVDESC;Catch_Net"
Dim pGP4 As IGeoProcessor
Set pGP4 = New GeoProcessor
Set pGPResults = Nothing
Set pGPResults = pGP4.Execute("Statistics_analysis", pParams, Nothing) 'here

Set pGPResults = Nothing
Set pGP4 = Nothing
Set pParams = Nothing

'#######################################################################################
'Create the scenario - writing the above summary info to the CatchP_S table rather than the standard values
'use the cmdCreateScenario function as we want to create a whole scenario - something the user can reload
'however use a Boolean to control loading from the summary table just created rather than other data
'#######################################################################################
Label37.Caption = "Finalising the scenario data and tidying up..."
DoEvents

blnUseModifiedLandCoverSlope = True
Set pTabColl = pMap 'refreshing this because of added tables
Dim intTempCounter As Integer
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = "summary_table" Then
        Set pTableUserDefinedLCoverSlope_Summary = pTabColl.StandaloneTable(intTempCounter)
    End If
Next
cmdCreateScenario_Click
blnUseModifiedLandCoverSlope = False    'set it back to false for general usage - this is only required within this command

'remove temporary layers to tidy up
Dim pFLayerToDelete As IFeatureLayer
FindFLayer pFLayerToDelete, "LocalCatchment_and_Network_Selected", True
pMap.DeleteLayer pFLayerToDelete

FindFLayer pFLayerToDelete, "CatchClipSlpClsLandCover", True
pMap.DeleteLayer pFLayerToDelete

FindFLayer pFLayerToDelete, "updated_baseline_landcover", True
pMap.DeleteLayer pFLayerToDelete

FindFLayer pFLayerToDelete, "CatchClipSlpClsLandCover_User_Inter", True
pMap.DeleteLayer pFLayerToDelete

pTabColl.RemoveStandaloneTable pTableUserDefinedLCoverSlope_Summary

If optDeleteIntermediateFiles Then
    fso.deletefile "C:\temp\CatchClipSlpClsLandCover.*", True
    fso.deletefile "C:\temp\updated_baseline_landcover.*", True
    fso.deletefile "C:\temp\summary_table.*", True
    fso.deletefile "C:\temp\CatchClipSlpClsLandCover_User_Inter.*", True
End If

'attempt to deal with Geoprocessing crashes
Set pGpUtils = New GPUtilities
pGpUtils.ReleaseInternals

Label37.Caption = "Done!"
DoEvents

DoEvents
Exit Sub
ErrorHandler: MsgBox "Unfortunately your processing has terminated due to an internal error or limitation in ArcGIS." & vbCrLf & "If the error persists please exit and restart ArcGIS.", vbCritical
Resume GetOut
GetOut:
Exit Sub

End Sub
Private Sub cmdLoadScenario_Click()
'#######################################################################################
'Load from the various tables the scenario corresponding to cboScenarioID, cboFilterScenarioName, cboFilterScenarioOwner, cboFilterScenarioDate,
'cboFilterScenarioComment - it is cboScenarioID that is the control
'#######################################################################################
CreateGBLakesTable
GetScenarioPointSourceIndices
blnScenarioLoaded = True

'populate cboGBLAKES_IDs with only the matching GBLAKES_IDs from cboScenarioID, within that scenario any of the included catchments can be processed
'the table in LocalCatchment_and_Network_S contains information for populating cboGBLAKES_IDs - use strScenarioLocalCatchmentAndNetwork and get its table
'find the scenario layer
FindFLayer pFLayerCatchment_Scenario, cboCatchmentScenario.Text, False
Set pFClassCatchment_Scenario = pFLayerCatchment_Scenario.FeatureClass
Set pScenarioLocalCatchmentAndNetworkTable = pFClassCatchment_Scenario

blnDataLoadedFromAScenario = True 'this is used in the creation of a scenario from a scenario
cboGBLAKES_IDs.Clear 'clear the existing combo

'ensure no duplicate data is created
Erase lonGBLAKES_IDNetworkMatchArray()
Erase lonOrderMatchArray()

Dim pQueryFilt As IQueryFilter2
Dim iOIDList() As Long
Dim lonRowScroller As Long
Dim pRow As IRow
Dim pCursor As ICursor
Set pQueryFilt = New QueryFilter    'don't give it any value yet
lonSelectedScenario = cboScenarioID.Value
pQueryFilt.WhereClause = "ScenarioID = " & cboScenarioID.Value

Set pCursor = pScenarioLocalCatchmentAndNetworkTable.Search(pQueryFilt, False)
lonRowScroller = pScenarioLocalCatchmentAndNetworkTable.RowCount(pQueryFilt)
ReDim lonGBLAKES_IDArray(lonRowScroller)
ReDim strSitenameArray(lonRowScroller)
ReDim intOrderArray(lonRowScroller)
ReDim lonNetworkArray(lonRowScroller)
lonRowScroller = 0
Set pRow = pCursor.NextRow
While Not pRow Is Nothing
'populate the combobox
    lonGBLAKES_IDArray(lonRowScroller) = pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("GBLAKES_ID"))
    strSitenameArray(lonRowScroller) = pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("SiteName"))
    intOrderArray(lonRowScroller) = pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("Order_"))
    lonNetworkArray(lonRowScroller) = pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("NETWORK"))
    cboGBLAKES_IDs.AddItem pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("GBLAKES_ID")) _
        & " - " & Translation(pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("SiteName")))
    If ReturnWFD_WB_ID(CLng(pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("GBLAKES_ID")))) <> 0 Then
        cboWBID.AddItem ReturnWFD_WB_ID(CLng(pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("GBLAKES_ID")))) _
             & " - " & Translation(pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("SiteName")))
    End If
    Set pRow = pCursor.NextRow
    lonRowScroller = lonRowScroller + 1
Wend

Set pCursor = pScenarioLocalCatchmentAndNetworkTable.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
cboGBLAKES_IDs.Value = pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("GBLAKES_ID")) & " - " & Translation(pRow.Value(pScenarioLocalCatchmentAndNetworkTable.FindField("SiteName")))

'from cmdLoad_Data
'get the selected flow table and read into an array - note that this is not a scenario flow routing table - there isn't one
Dim intTempCounter As Integer
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblFlowRouting Then
        Set pFlowRoutingTable = pTabColl.StandaloneTable(intTempCounter)
    End If
Next

Dim intRouteGBLAKES_IDField As Integer
Dim intDSGBLAKES_IDField As Integer
Dim lonRowScroller2 As Long
Dim pQueryFilt2 As IQueryFilter2
Set pQueryFilt2 = New QueryFilter    'don't give it any value yet
lonFlowRoutes = pFlowRoutingTable.Table.RowCount(pQueryFilt2)

Erase arrayFlowRouting()
ReDim arrayFlowRouting(lonFlowRoutes, 1)
Dim lonFlowRowScroller2 As Long
Dim iOIDFlowList() As Long
If lonFlowRoutes > 1 Then
    ReDim iOIDFlowList(lonFlowRoutes - 1)
Else
    ReDim iOIDFlowList(lonFlowRoutes)
End If

For lonRowScroller2 = 0 To lonFlowRoutes - 1
    iOIDFlowList(lonRowScroller2) = lonRowScroller2 + 1 'add the plus + as the first row isn't zero
Next

intRouteGBLAKES_IDField = pFlowRoutingTable.Table.FindField("GBLAKES_ID")
If intRouteGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID", vbCritical
    Exit Sub
End If
intDSGBLAKES_IDField = pFlowRoutingTable.Table.FindField("DownstreamGBLAKES_ID")
If intDSGBLAKES_IDField = -1 Then
    MsgBox "Cannot find field DS_SITECOD", vbCritical
    Exit Sub
End If

Set pCursor = Nothing
Set pCursor = pFlowRoutingTable.Table.GetRows(iOIDFlowList, True)

If Not pFlowRoutingTable.Table.HasOID Then
    MsgBox pFlowRoutingTable.Name & " is not registered. In ArcCatalog right click the table and choose Register with GDB, do the same with all tables.", vbCritical
    Exit Sub
End If

Set pRow = pCursor.NextRow
lonRowScroller2 = 0
While Not pRow Is Nothing
'populate the arrays that we will use for subsequent processing
    arrayFlowRouting(lonRowScroller2, 0) = pRow.Value(intRouteGBLAKES_IDField)
    arrayFlowRouting(lonRowScroller2, 1) = pRow.Value(intDSGBLAKES_IDField)
    Set pRow = pCursor.NextRow
    lonRowScroller2 = lonRowScroller2 + 1
Wend

'enable the forms
cmdGetCatchmentInfo.Enabled = True
cmdZoomToSelected.Enabled = True
cmdHighlightSelected.Enabled = True
cmdIntersectUserShapefile.Enabled = True
cboGBLAKES_IDs.Enabled = True
cboWBID.Enabled = True
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label49.Enabled = True
Label50.Enabled = True
Label51.Enabled = True
optCatchment.Enabled = True
lblScenarioSaveWarning.Visible = True
lblScenarioSaveWarning2.Visible = True
lblReportSaveWarning.Visible = True
Frame1Scenario.Visible = False
Frame2UserLandCover.Visible = False
cmdZoomToSelected.Enabled = False
cmdHighlightSelected.Enabled = False
cmdCreateReport.Enabled = False
DoEvents

LoadSepaMonitoringIntoArray

End Sub

Private Sub cmdModCatchmentInputs_Click()
'Note that there are limitations in changing the P of a user chosen land use to the same value for the whole network - these
'individual P values are slope specific, so each catchment varies... so the accurate approach is to use the SlopeClass_LandCover.shp

'#######################################################################################
'Procedure:
'Implement the user entered P or area values
'get selections:
'change P of selected cover for selected catchment
'change P of selected cover for whole network
'change area of selected land cover
'#######################################################################################
'the input data to modify comes from tblCatchP
'we have an array of the inputs for the selected catchment
Dim m As Integer
'Get user entered P
If (chkChangeP Or chkChangePforNetwork) And optUserModified Then
    If IsNumeric(txtEnterNewP) Then
        dblUser_Entered_P_for_Selected_Site = CDbl(txtEnterNewP)
        blnNoLongerBaseline = True
        'look for an existing instance of the user selected land cover and over-write it or find the first un-set location
        For m = 0 To UBound(varModifiedLCoverCoeff, 1)
            If varModifiedLCoverCoeff(m, 0) = cboResolveAreaDifference.Value Then
                varModifiedLCoverCoeff(m, 1) = CDbl(txtEnterNewP.Text)
                Exit For
            End If
            If varModifiedLCoverCoeff(m, 0) = "" Then
                varModifiedLCoverCoeff(m, 0) = cboResolveAreaDifference.Value
                varModifiedLCoverCoeff(m, 1) = CDbl(txtEnterNewP.Text)
                Exit For
            End If
        Next
    Else
        blnNoLongerBaseline = False
        MsgBox "Your value for P is not a number", vbCritical
        Exit Sub
    End If
End If

'#######################################################################################
'Get user entered new area
'#######################################################################################
If chkChangeArea Then
    If IsNumeric(txtEnterNewArea) Then
        dblUser_Entered_Area_for_Selected_Site = CDbl(txtEnterNewArea)
        blnNoLongerBaseline = True
    Else
        MsgBox "Your value for area is not a number", vbCritical
        blnNoLongerBaseline = False
        Exit Sub
    End If
End If

'#######################################################################################
'Check if the area and revised area are both zero, if they are delete the records
'#######################################################################################
Dim intNonZeroCounter As Integer
intNonZeroCounter = 0
Dim i As Integer
For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Or arrayCatchPforChosenGBLAKES_ID(i, 5) > 0 Then
        intNonZeroCounter = intNonZeroCounter + 1
    End If
Next
Dim j As Integer
j = 0
If intNonZeroCounter <= UBound(arrayCatchPforChosenGBLAKES_ID, 1) Then
    'there are zero land covers so remove them from the array
    ReDim varTempArray(intNonZeroCounter - 1, 7)
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
         If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Or arrayCatchPforChosenGBLAKES_ID(i, 5) > 0 Then
            varTempArray(j, 0) = arrayCatchPforChosenGBLAKES_ID(i, 0)
            varTempArray(j, 1) = arrayCatchPforChosenGBLAKES_ID(i, 1)
            varTempArray(j, 2) = arrayCatchPforChosenGBLAKES_ID(i, 2)
            varTempArray(j, 3) = arrayCatchPforChosenGBLAKES_ID(i, 3)
            varTempArray(j, 4) = arrayCatchPforChosenGBLAKES_ID(i, 4)
            varTempArray(j, 5) = arrayCatchPforChosenGBLAKES_ID(i, 5)
            varTempArray(j, 6) = arrayCatchPforChosenGBLAKES_ID(i, 6)
            varTempArray(j, 7) = arrayCatchPforChosenGBLAKES_ID(i, 7)
            j = j + 1
         End If
    Next
    ReDim arrayCatchPforChosenGBLAKES_ID(intNonZeroCounter - 1, 7)
    For i = 0 To UBound(varTempArray, 1)
        arrayCatchPforChosenGBLAKES_ID(i, 0) = varTempArray(i, 0)
        arrayCatchPforChosenGBLAKES_ID(i, 1) = varTempArray(i, 1)
        arrayCatchPforChosenGBLAKES_ID(i, 2) = varTempArray(i, 2)
        arrayCatchPforChosenGBLAKES_ID(i, 3) = varTempArray(i, 3)
        arrayCatchPforChosenGBLAKES_ID(i, 4) = varTempArray(i, 4)
        arrayCatchPforChosenGBLAKES_ID(i, 5) = varTempArray(i, 5)
        arrayCatchPforChosenGBLAKES_ID(i, 6) = varTempArray(i, 6)
        arrayCatchPforChosenGBLAKES_ID(i, 7) = varTempArray(i, 7)
    Next
End If

'run the calculation
CalculateTP

If blnNoLongerBaseline Then
    SwitchReportToScenario
Else
    SwitchReportToBaseline
End If
End Sub
Private Sub cmdResetChanges_Click()
txtPerCapitaTPLoadUrban.Text = ""
txtPerCapitaTPLoadRural.Text = ""
txtUrbanPop.Text = ""
txtRuralPop.Text = ""
txtPointSourceAmount.Text = ""

chkPerCapitaTPLoadUrban.Value = False
chkPerCapitaTPLoadRural.Value = False
chkPerCapitaTPLoadUrbanAll.Value = False
chkPerCapitaTPLoadRuralAll.Value = False

chkUrbanPop.Value = False
chkRuralPop.Value = False
chkAddPointSource = False
chkRemoveSelectedPointSources = False
blnModifyOtherPointSourceLoad = False
blnModifySewageLoad = False
Erase varPointSource()
If tglMasterOrScenario Then 'we are running a scenario so need to reload the point sources (they have just been erased)
    cmdGetCatchmentInfo_Click
    CalculateTP
Else
    CalculateTP
End If
CalculateTP
SwitchReportToBaseline

End Sub
Private Sub cmdResetModifiedValues_Click()
'#######################################################################################
'Reset the user modifications to P and area
'#######################################################################################

'reset 'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
'i.e. copy the read in values to the modified
Dim Response As String
Dim i As Integer
Response = MsgBox("Are you sure you wish to revert to read-in values?", vbYesNo)
If Response = vbYes Then
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
        arrayCatchPforChosenGBLAKES_ID(i, 5) = arrayCatchPforChosenGBLAKES_ID(i, 3) 'this will contain modified area
        arrayCatchPforChosenGBLAKES_ID(i, 6) = arrayCatchPforChosenGBLAKES_ID(i, 4) 'this will contain modified kg/ha
        arrayCatchPforChosenGBLAKES_ID(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 2) 'this will contain revised P
    Next
    txtEnterNewArea = ""
    txtEnterNewP = ""
    chkChangeArea = False
    chkChangeP = False
    chkChangePforNetwork = False
End If
optReadIn = True
CalculateTP

SwitchReportToBaseline
blnNoLongerBaseline = False

'clear the array of modified land covers - want to maintain the structure, so no erase
For i = 0 To UBound(varModifiedLCoverCoeff, 1)
    varModifiedLCoverCoeff(i, 0) = ""
    varModifiedLCoverCoeff(i, 1) = 0
Next
End Sub
Private Sub cmdResolveAreaDifference_Click()
'#######################################################################################
'This is to resolve any discrepancy in the sum of areas after a user modifies the area of a selected land cover,
'it uses lvwCatchmentInfo.SelectedItem.SubItems(2) - which is strSelectedLandCoverType
'to detect the correct land cover the area difference is in dblUserModifiedLCoverArea_difference - and this is signed - I hide the sign in the output
'modify arrayCatchPforChosenGBLAKES_ID(i, 5) where
'arrayCatchPforChosenGBLAKES_ID(i, 1) = lvwCatchmentInfo.SelectedItem.SubItems(2) also equals strSelectedLandCoverType
'#######################################################################################

If blnUseLCoverComboSelection Then
    strSelectedLandCoverType = cboResolveAreaDifference.Value
End If

'check if the user selected land cover already exists for the chosen catchment while calculating the area to resolve
Dim blnSelectedLandCoverAlreadyExists As Boolean
blnSelectedLandCoverAlreadyExists = False
blnResolveDifferences = True
Dim i As Integer
For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    If arrayCatchPforChosenGBLAKES_ID(i, 1) = strSelectedLandCoverType Then
        blnSelectedLandCoverAlreadyExists = True
        'check in case the user tries to subtract too much...
        If Abs(dblUserModifiedLCoverArea_difference) > Abs(arrayCatchPforChosenGBLAKES_ID(i, 5)) And dblUserModifiedLCoverArea_difference > arrayCatchPforChosenGBLAKES_ID(i, 5) Then
            arrayCatchPforChosenGBLAKES_ID(i, 5) = 0
        Else
            arrayCatchPforChosenGBLAKES_ID(i, 5) = arrayCatchPforChosenGBLAKES_ID(i, 5) - dblUserModifiedLCoverArea_difference
        End If
        arrayCatchPforChosenGBLAKES_ID(i, 7) = arrayCatchPforChosenGBLAKES_ID(i, 6) * (arrayCatchPforChosenGBLAKES_ID(i, 5) / 10000)
    End If
Next

'certain combinations of buttons may result in txtEnterNewP.Text being "" and cmdResolveAreaDifference.caption having a value
'so use the cmdResolveAreaDifference.caption to populate txtEnterNewP.Text
Dim strTempText As String
If txtEnterNewP.Text = "" And Len(cmdResolveAreaDifference.Caption) > 37 Then
    strTempText = Left(cmdResolveAreaDifference.Caption, Len(cmdResolveAreaDifference.Caption) - 7)
    cboResolveAreaDifference.Value = Right(strTempText, Len(strTempText) - 37)
End If

'add the new land cover if not found above
'need to enlarge the multidimensional array by copying it then redim without a preserve
If Not blnSelectedLandCoverAlreadyExists Then
    Dim varTempArray() As Variant
    Dim j As Integer
    ReDim varTempArray(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 7)
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
        For j = 0 To 7
            varTempArray(i, j) = arrayCatchPforChosenGBLAKES_ID(i, j)
        Next j
    Next
    
    i = UBound(arrayCatchPforChosenGBLAKES_ID, 1) + 1
    ReDim arrayCatchPforChosenGBLAKES_ID(i, 7)
    
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1) - 1 'minus one because varTempArray contains only the copied data, not the additional record
        For j = 0 To 7
            arrayCatchPforChosenGBLAKES_ID(i, j) = varTempArray(i, j)
        Next j
    Next
    If txtEnterNewP.Text = "" Then
        MsgBox "Warning, inappropriate selection made, cannot process.", vbCritical
        Exit Sub
    End If
    'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 0) = lonChosenGBLAKES_ID
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 1) = cboResolveAreaDifference.Value & " (user)"
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 2) = 0    'there was no original value so must be 0 (and this is used in scenario creation)
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 3) = 0
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 4) = 0
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 5) = CDbl(txtEnterNewArea.Text)
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 6) = CDbl(txtEnterNewP.Text)
    arrayCatchPforChosenGBLAKES_ID(UBound(arrayCatchPforChosenGBLAKES_ID, 1), 7) = CDbl(txtEnterNewP.Text) * CDbl(txtEnterNewArea.Text) / 10000
End If

'check if the area and revised area are both zero, if they are delete the records
Dim intNonZeroCounter As Integer
intNonZeroCounter = 0
For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
    If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Or arrayCatchPforChosenGBLAKES_ID(i, 5) > 0 Then
        intNonZeroCounter = intNonZeroCounter + 1
    End If
Next
j = 0
If intNonZeroCounter <= UBound(arrayCatchPforChosenGBLAKES_ID, 1) Then
    'there are zero land covers so remove them from the array
    ReDim varTempArray(intNonZeroCounter - 1, 7)
    For i = 0 To UBound(arrayCatchPforChosenGBLAKES_ID, 1)
         If arrayCatchPforChosenGBLAKES_ID(i, 3) > 0 Or arrayCatchPforChosenGBLAKES_ID(i, 5) > 0 Then
            varTempArray(j, 0) = arrayCatchPforChosenGBLAKES_ID(i, 0)
            varTempArray(j, 1) = arrayCatchPforChosenGBLAKES_ID(i, 1)
            varTempArray(j, 2) = arrayCatchPforChosenGBLAKES_ID(i, 2)
            varTempArray(j, 3) = arrayCatchPforChosenGBLAKES_ID(i, 3)
            varTempArray(j, 4) = arrayCatchPforChosenGBLAKES_ID(i, 4)
            varTempArray(j, 5) = arrayCatchPforChosenGBLAKES_ID(i, 5)
            varTempArray(j, 6) = arrayCatchPforChosenGBLAKES_ID(i, 6)
            varTempArray(j, 7) = arrayCatchPforChosenGBLAKES_ID(i, 7)
            j = j + 1
         End If
    Next
    ReDim arrayCatchPforChosenGBLAKES_ID(intNonZeroCounter - 1, 7)
    For i = 0 To UBound(varTempArray, 1)
        arrayCatchPforChosenGBLAKES_ID(i, 0) = varTempArray(i, 0)
        arrayCatchPforChosenGBLAKES_ID(i, 1) = varTempArray(i, 1)
        arrayCatchPforChosenGBLAKES_ID(i, 2) = varTempArray(i, 2)
        arrayCatchPforChosenGBLAKES_ID(i, 3) = varTempArray(i, 3)
        arrayCatchPforChosenGBLAKES_ID(i, 4) = varTempArray(i, 4)
        arrayCatchPforChosenGBLAKES_ID(i, 5) = varTempArray(i, 5)
        arrayCatchPforChosenGBLAKES_ID(i, 6) = varTempArray(i, 6)
        arrayCatchPforChosenGBLAKES_ID(i, 7) = varTempArray(i, 7)
    Next
End If

'tidy up
txtEnterNewArea.Text = ""

CalculateTP

End Sub
Private Sub cmdSewageApply_Click()
Dim i As Integer
Dim j As Integer
Dim blnArrayIsNotEmpty As Boolean
blnArrayIsNotEmpty = False
blnModifySewageLoad = True
If chkAddPointSource Then
    blnModifyOtherPointSourceLoad = True
    If Not chkRemoveSelectedPointSources Then   'this should not really be needed, just a precaution.
        'format is GBLAKES_ID, source type, source amount, scenario id
        'in this configuration of check boxes we are always adding, so redim the long way (can't preserve multi-D)
        On Error Resume Next
        blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
        If blnArrayIsNotEmpty Then  'add a row
            Dim varTempArray() As Variant
            varTempArray() = varPointSource()
            Erase varPointSource()
            ReDim varPointSource(UBound(varTempArray, 1) + 1, 2)
            For i = 0 To UBound(varTempArray, 1)
                varPointSource(i, 0) = varTempArray(i, 0)
                varPointSource(i, 1) = varTempArray(i, 1)
                varPointSource(i, 2) = varTempArray(i, 2)
            Next
            Erase varTempArray()
        Else    'dimension the array for the first time to receive the first points
            ReDim varPointSource(0, 2)
        End If
        varPointSource(UBound(varPointSource, 1), 0) = lonChosenGBLAKES_ID
        varPointSource(UBound(varPointSource, 1), 1) = cboPointSourceType.Value
        varPointSource(UBound(varPointSource, 1), 2) = CDbl(txtPointSourceAmount.Text)
    End If
Else
    blnModifyOtherPointSourceLoad = False
End If
If chkRemoveSelectedPointSources Then
'check for attempt to remove from an empty array
blnArrayIsNotEmpty = False
On Error Resume Next
blnArrayIsNotEmpty = UBound(varPointSource, 1) > -1
    If blnArrayIsNotEmpty Then
    'remove highlighted point sources in  from the list
    'match strSelectedPointSourceForRemoval with the item in the array
        For i = 0 To UBound(varPointSource, 1)
            If varPointSource(i, 1) = strSelectedPointSourceForRemoval Then
            'remove it
                If i < UBound(varPointSource, 1) Then 'it's not the last one
                    For j = i To UBound(varPointSource, 1) - 1
                        varPointSource(j, 0) = varPointSource(j + 1, 0)
                        varPointSource(j, 1) = varPointSource(j + 1, 1)
                        varPointSource(j, 2) = varPointSource(j + 1, 2)
                    Next
                End If
                If UBound(varPointSource, 1) = 0 Then 'empty the array - there was only one point in it
                    Erase varPointSource()
                    Exit For
                End If
                'delete the last set of records
                ReDim varTempArray(UBound(varPointSource, 1) - 1, 2)
                For j = 0 To UBound(varTempArray, 1)
                    varTempArray(j, 0) = varPointSource(j, 0)
                    varTempArray(j, 1) = varPointSource(j, 1)
                    varTempArray(j, 2) = varPointSource(j, 2)
                Next
                Erase varPointSource()
                varPointSource() = varTempArray()
                Erase varTempArray()
                Exit For
            End If
        Next
        'this is here as a precaution
        blnModifyOtherPointSourceLoad = True
    End If
End If

CalculateTP

End Sub
Private Sub cmdZoomToSelected_Click()
'#######################################################################################
'Zoom to the catchment selected in cboGBLAKES_IDs using the layer in cboCatchment - don't use earlier found layer
'   (in case the selection has been redone)
'get the selected catchment and flow routing tables from the cbos
'find the catchment layer
'#######################################################################################

'Switch to data view
If pMxDoc.ActiveView Is pMxDoc.PageLayout Then
    Set pMxDoc.ActiveView = pMxDoc.Maps.Item(0)
End If

FindFLayer pFLayerCatchment, cboCatchment.Text, False
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView
Dim varSplit As Variant

If cboGBLAKES_IDs.Text <> "" Then
    varSplit = Split(cboGBLAKES_IDs.Text, " - ")
    lonChosenGBLAKES_ID = varSplit(0)
    strChosenSitename = varSplit(1)
End If

Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment
If optCatchment Then
    pQueryFilter.WhereClause = "GBLAKES_ID =" & lonChosenGBLAKES_ID
Else
'zoom to network lonChosenNetwork
    pQueryFilter.WhereClause = "NETWORK =" & lonChosenNetwork
End If

'Perform the selection
pFeatureselection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
'Flag the new selection
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing

'zoom
Dim pSelSet As ISelectionSet
Set pSelSet = pFeatureselection.SelectionSet

Dim pEnumGeom As IEnumGeometry
Dim pEnumGeomBind As IEnumGeometryBind

Set pEnumGeom = New EnumFeatureGeometry
Set pEnumGeomBind = pEnumGeom
pEnumGeomBind.BindGeometrySource Nothing, pSelSet

Dim pGeomFactory As IGeometryFactory
Set pGeomFactory = New GeometryEnvironment

Dim pGeom As IGeometry
Set pGeom = pGeomFactory.CreateGeometryFromEnumerator(pEnumGeom)

pMxDoc.ActiveView.Extent = pGeom.Envelope
pMxDoc.ActiveView.Refresh

End Sub

Private Sub txtOutputFile_Change()
    If chkProduceResultsCSV.Enabled Then chkProduceResultsCSV.Value = True
End Sub
Private Sub txtOutputReport_Change()
    If chkProduceResultsTable.Enabled Then chkProduceResultsTable.Value = True
End Sub
Private Sub txtPointSourceAmount_Change()
If txtPointSourceAmount.Value <> "" Then
    chkAddPointSource.Value = True
    SwitchReportToScenario
Else
    chkAddPointSource.Value = False
    SwitchReportToBaseline
End If
End Sub
Private Sub lvwCatchmentInfo_Click()
'#######################################################################################
'Allow the user to modify P or Area by selecting a land cover
'#######################################################################################
Dim intIndexSelectedLandCover As Integer
Dim i As Integer
Dim dblP As Double
Dim dblArea As Double
Dim dblPperArea As Double
lvwCatchmentInfo.HideSelection = False
Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter

'check if the user has selected one of the user added land covers, need to remove the ' (user)' or ' (mod.)text if they have
Dim strSelectedItem As String
'note that " - 7" is to remove the characters like (kg) etc that I add for the display
If lvwCatchmentInfo.SelectedItem.SubItems(2) Like "* (user)" Then
    strSelectedItem = Left(lvwCatchmentInfo.SelectedItem.SubItems(2), Len(lvwCatchmentInfo.SelectedItem.SubItems(2)) - 7)
Else
    strSelectedItem = lvwCatchmentInfo.SelectedItem.SubItems(2)
End If
If lvwCatchmentInfo.SelectedItem.SubItems(2) Like "* (mod.)" Then
    strSelectedItem = Left(lvwCatchmentInfo.SelectedItem.SubItems(2), Len(lvwCatchmentInfo.SelectedItem.SubItems(2)) - 7)
Else
    strSelectedItem = lvwCatchmentInfo.SelectedItem.SubItems(2)
End If
cboResolveAreaDifference.Value = strSelectedItem

If lvwCatchmentInfo.SelectedItem.Text <> "" And Len(lvwCatchmentInfo.SelectedItem.Text) > 7 Then
    strSelectedPointSourceForRemoval = Left(lvwCatchmentInfo.SelectedItem.Text, Len(lvwCatchmentInfo.SelectedItem.Text) - 7)
Else
    strSelectedPointSourceForRemoval = ""
End If

pQueryFilt.WhereClause = "LCOVDESC = " & Chr(34) & strSelectedItem & Chr(34)
i = pExportsTable.Table.RowCount(pQueryFilt)

If i > 0 Then
        lblSelectedLandCover.Font.size = 8  'this is changing on the display so fixing it here
        'if the area is set to zero
        dblP = lvwCatchmentInfo.SelectedItem.SubItems(1)
        If lvwCatchmentInfo.SelectedItem.SubItems(3) = 0 Then
            dblArea = 0
            dblPperArea = 0
        Else
            dblArea = lvwCatchmentInfo.SelectedItem.SubItems(3) / 10000 'convert metres -> hectares
            dblPperArea = dblP / dblArea
        End If
        lblSelectedLandCover = lvwCatchmentInfo.SelectedItem.SubItems(2) & ", P = " & Format(dblPperArea, "#0.000") & " kg/ha"
        txtEnterNewP.Value = Format(dblPperArea, "#0.000")
        chkChangeP.Caption = "Change P export coeff. of " & lvwCatchmentInfo.SelectedItem.SubItems(2)
        chkChangeArea.Caption = "Change Area of " & lvwCatchmentInfo.SelectedItem.SubItems(2)
        chkChangePforNetwork.Caption = "Change P export coeff. for " & lvwCatchmentInfo.SelectedItem.SubItems(2) & " for whole network"
        strLcoverForNetworkChange = lvwCatchmentInfo.SelectedItem.SubItems(2)
        txtEnterNewP.Visible = True
        txtEnterNewArea.Visible = True
        txtEnterNewP.Enabled = True
        txtEnterNewArea.Enabled = True
        optReadIn.Visible = True
        optUserModified.Visible = True
        lblEnterNewP.Caption = "Enter new P (kg/ha)"
        lblEnterNewArea.Caption = "Enter new area (m" & Chr(178) & ")"
        cmdModCatchmentInputs.Enabled = True
        cmdResetModifiedValues.Enabled = True
        chkChangeP.Enabled = True
        chkChangeArea.Enabled = True
        cmdInfoChangeP.Enabled = True
        chkChangePforNetwork.Enabled = True
        cmdResolveAreaDifference.Caption = "Resolve area difference by adjusting " & lvwCatchmentInfo.SelectedItem.SubItems(2)
        DoEvents
    Else
        lblSelectedLandCover = ""
        txtEnterNewP.Value = ""
        chkChangeP.Caption = "Change P export coeff."
        chkChangeArea.Caption = "Change Area"
        chkChangePforNetwork.Caption = "Change P export coeff. for whole network"
        txtEnterNewP.Visible = False
        txtEnterNewArea.Visible = False
        optReadIn.Visible = False
        optUserModified.Visible = False
        lblEnterNewP.Caption = ""
        lblEnterNewArea.Caption = ""
        cmdModCatchmentInputs.Enabled = False
        cmdResetModifiedValues.Enabled = False
        chkChangeP.Enabled = False
        chkChangeArea.Enabled = False
        cmdInfoChangeP.Enabled = False
        chkChangePforNetwork.Enabled = False
        cmdResolveAreaDifference.Caption = "Resolve area difference"
        DoEvents
End If

strSelectedLandCoverType = lvwCatchmentInfo.SelectedItem.SubItems(2)
'as this list view has been clicked it becomes the priority
blnUseLCoverComboSelection = False
DoEvents
End Sub
Private Sub lvwCatchmentRelationships1_Click()
'#######################################################################################
'Highlight the clicked feature on the map
'Highlight the catchment selected in cboGBLAKES_IDs using the layer in cboCatchment - don't use earlier found layer
'in case the selection has been redone
'get the selected catchment and flow routing tables from the cbos
'find the catchment layer
'#######################################################################################
FindFLayer pFLayerCatchment, cboCatchment.Text, False
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView

'name to process is in lvwCatchmentRelationships1.SelectedItem.Text
Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment
pQueryFilter.WhereClause = "GBLAKES_ID =" & lvwCatchmentRelationships1.SelectedItem.Text

'Invalidate only the selection cache
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing 'Flag the original selection
pFeatureselection.SelectFeatures pQueryFilter, esriSelectionResultNew, False 'Perform the selection
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing 'Flag the new selection

End Sub
Private Sub lvwCatchmentRelationships2_Click()
'#######################################################################################
'Highlight the clicked feature on the map
'highlight the catchment selected in cboGBLAKES_IDs using the layer in cboCatchment - don't use earlier found layer
'in case the selection has been redone
'get the selected catchment and flow routing tables from the cbos
'find the catchment layer
'#######################################################################################
FindFLayer pFLayerCatchment, cboCatchment.Text, False
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView

'name to process is in lvwCatchmentRelationships1.SelectedItem.Text
Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment
pQueryFilter.WhereClause = "GBLAKES_ID =" & lvwCatchmentRelationships2.SelectedItem.SubItems(2)

'Invalidate only the selection cache
'Flag the original selection
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
pFeatureselection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing

End Sub
Private Sub optCatchment_Click()
cmdZoomToSelected.Caption = "Zoom to catchment " & lonChosenGBLAKES_ID
cmdHighlightSelected.Caption = "Highlight catchment " & lonChosenGBLAKES_ID
End Sub
Private Sub optJPEG_Click()
txtOutputRatio.Enabled = False
SpinButtonSampleRatio.Enabled = False
Label43.Enabled = False
txtJPEGQuality.Enabled = True
SpinButtonJPEGQuality.Enabled = True
Label44.Enabled = True
End Sub
Private Sub optModifiedData_Click()
If optOriginalData Then
    txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using original base line data."
Else
    txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using scenario data."
End If
End Sub
Private Sub optNetwork_Click()
cmdZoomToSelected.Caption = "Zoom to selected network"
cmdHighlightSelected.Caption = "Highlight selected network"
End Sub
Private Sub optOriginalData_Click()
If optOriginalData Then
    txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using original base line data."
Else
    txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using scenario data."
End If
End Sub
Private Sub optPDF_Click()
txtOutputRatio.Enabled = True
SpinButtonSampleRatio.Enabled = True
Label43.Enabled = True
txtJPEGQuality.Enabled = False
SpinButtonJPEGQuality.Enabled = False
Label44.Enabled = False
End Sub
Private Sub SpinButtonJPEGQuality_Change()
    txtJPEGQuality = SpinButtonJPEGQuality.Value
End Sub
Private Sub SpinButtonSampleRatio_Change()
    txtOutputRatio = SpinButtonSampleRatio.Value
End Sub
Private Sub tglMasterOrScenario_Click()
Run_MasterOrScenario

SwitchReportToBaseline
blnNoLongerBaseline = False
End Sub
Sub Run_MasterOrScenario()
'#######################################################################################
'Respond to the click of the toggle buttons by changing the form contents
'#######################################################################################
Dim i As Integer
Dim pCursor As ICursor
Dim pRow As IRow
cboGBLAKES_IDs.Clear
cboWBID.Clear
If tglMasterOrScenario Then
    cboScenarioID.Clear
    cboFilterScenarioName.Clear
    cboFilterScenarioOwner.Clear
    cboFilterScenarioDate.Clear
    cboFilterScenarioComment.Clear
    tglMasterOrScenario.Caption = "Click to load from master data"
    frameLoadFromMaster.Visible = False
    frameLoadFromScenario.Visible = True
    'repopulate the pull down menus
    'check that the ScenarioID is new  or offer the user the opportunity to choose another
    For i = 0 To pTabColl.StandaloneTableCount - 1
        If pTabColl.StandaloneTable(i).Name = strScenarioTbl Then
            Set pScenario = pTabColl.StandaloneTable(i)
        End If
    Next
    intScenarioIDField = pScenario.Table.FindField("ScenarioID")
    intScenarioNameField = pScenario.Table.FindField("ScenarioName")
    intScenarioCreatorField = pScenario.Table.FindField("ScenarioCreator")
    intScenarioCreationDateField = pScenario.Table.FindField("ScenarioCreationDate")
    intScenarioCommentField = pScenario.Table.FindField("Comment")
    intScenarioRegionField = pScenario.Table.FindField("Region")
    If intScenarioIDField = -1 Then
        MsgBox "Cannot find field ScenarioID", vbCritical
        Exit Sub
    End If
    Dim pQueryFilt As IQueryFilter2
    Set pQueryFilt = New QueryFilter
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    While Not pRow Is Nothing
        cboScenarioID.AddItem pRow.Value(intScenarioIDField)
        cboFilterScenarioName.AddItem pRow.Value(intScenarioNameField)
        cboFilterScenarioOwner.AddItem pRow.Value(intScenarioCreatorField)
        cboFilterScenarioDate.AddItem Day(pRow.Value(intScenarioCreationDateField)) & "/" & Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
        cboFilterScenarioComment.AddItem pRow.Value(intScenarioCommentField)
        Set pRow = pCursor.NextRow
    Wend
    'once again so that the value in the cboboxes is the first found
    Set pCursor = pScenario.Table.Search(pQueryFilt, False)
    Set pRow = pCursor.NextRow
    If pScenario.Table.RowCount(pQueryFilt) = 0 Then
        MsgBox "Your scenario table is empty, therefore there is nothing to load from it." _
        & vbCrLf & "Please create a scenario prior to attempting to load one.", vbCritical, "Invalid entry"
        Exit Sub
    End If
    cboScenarioID.Value = pRow.Value(intScenarioIDField)
    cboFilterScenarioName.Value = pRow.Value(intScenarioNameField)
    cboFilterScenarioOwner.Value = pRow.Value(intScenarioCreatorField)
    cboFilterScenarioDate.Value = Day(pRow.Value(intScenarioCreationDateField)) & "/" & Month(pRow.Value(intScenarioCreationDateField)) & "/" & Year(pRow.Value(intScenarioCreationDateField))
    cboFilterScenarioComment.Value = pRow.Value(intScenarioCommentField)
Else
    tglMasterOrScenario.Caption = "Click to load from scenario"
    frameLoadFromMaster.Visible = True
    frameLoadFromScenario.Visible = False
End If
End Sub
Private Sub txtBoxOtherRegion_Change()
    optRegionN.Value = False
    optRegionSW.Value = False
    optRegionSE.Value = False
End Sub
Private Sub txtEnterNewArea_Change()
If txtEnterNewArea.Value <> "" Then
    chkChangeArea.Value = True
    Else
    chkChangeArea.Value = False
End If
End Sub
Private Sub txtEnterNewP_Change()
If txtEnterNewP.Value <> "" Then
    chkChangeP.Value = True
Else
    chkChangeP.Value = False
End If
End Sub
Private Sub txtPerCapitaTPLoadRural_Change()
If txtPerCapitaTPLoadRural.Value <> "" Then
    chkPerCapitaTPLoadRural.Value = True
    SwitchReportToScenario
Else
    chkPerCapitaTPLoadRural.Value = False
End If
End Sub
Private Sub txtPerCapitaTPLoadUrban_Change()
If txtPerCapitaTPLoadUrban.Value <> "" Then
    chkPerCapitaTPLoadUrban.Value = True
    SwitchReportToScenario
Else
    chkPerCapitaTPLoadUrban.Value = False
End If
End Sub
Private Sub txtRuralPop_Change()
If txtRuralPop.Value <> "" Then
    chkRuralPop.Value = True
    SwitchReportToScenario
Else
    chkRuralPop.Value = False
End If
End Sub
Private Sub txtScenarioID_2_Change()
txtScenarioID = txtScenarioID_2
DoEvents
End Sub
Private Sub txtUrbanPop_Change()
If txtUrbanPop.Value <> "" Then
    chkUrbanPop.Value = True
    SwitchReportToScenario
Else
    chkUrbanPop.Value = False
End If
End Sub
Private Sub UserForm_Activate()
'test the MSComctlLib library is loaded
blnNoLongerBaseline = False
'This test to check references doesn't work at the SEPA site, although it does work at MLURI/JHI
'retained in case a SEPA developer wants to take a look
'If Not CheckReferencesAttached Then
'    MsgBox "Warning, the Microsoft ImageList Control 6.0 library is not referenced in the Visual Basic environment." _
'    & vbCrLf & "This must be referenced for the tool to function. Please consult the user guide.", vbCritical
'    Exit Sub
'End If

'May want to alter this section to have a text configuration file controlling the various criteria in a future version
Set fso = CreateObject("scripting.filesystemobject")
Dim intTemp As Integer
Dim j As Long
Dim intTempCounter As Integer

DoEvents

intTopListviewBoxesTop = 55
intTopListviewBoxesHeight = 92
frameLoadFromScenario.Top = 22
frameLoadFromScenario.Left = 10
MultiPage1.Pages(4).Enabled = True
MultiPage1.Pages(4).Caption = "Additional results"

'initialise the environment
Set pMxDoc = ThisDocument
Set pMap = pMxDoc.FocusMap
Set pTabColl = pMap
Set pMxApp = Application

'set the boolean for scenario saving to false, it will be set to true if the conditions in InitialiseScenarioSave are true
blnScenarioCanBeSaved = False
blnDataLoadedFromAScenario = False

'prevents an out of range error:
ReDim pGBLakes_WBID_Array(1, 1)

'###################################################################################
'This section is for modelling Skene catchment analysis
'standard SEPA setting is:
'blnSkeneRun = False
'###################################################################################

'set the boolean for using user modified land cover data in a scenario save to false
'this will only be set to true if the 'Import user created land cover polygon(s)' tab triggers the scenario creation
'blnUseModifiedLandCoverSlope = True
Dim blnSkeneRun As Boolean
blnSkeneRun = False 'if False then it is not a scenario test
If (blnSkeneRun) Then
    MsgBox "This is a Skene modified PLUS+"
    'may want to put the following into a text file for the user to load as a configuration
    strTblCatchmentSewageName = "Skene_CatchmentSewage"
    'strTblCatchPName = "Merge_Baseline_CatchP_WEAG"  'this is the landcover sourced P for each catchment, modified with WEAG
    strTblCatchPName = "Merge_Baseline_CatchP_Summary"  'this is the landcover sourced P for each catchment
    strTblExportsName = "Exports"
    strTblFlowRouting = "Skene_FlowRouting"
    strTblLoadPrecursorName = "Skene_LoadPrecursor"
    strTblPerCapitaTPLoadsName = "PerCapitaTPLoads"
    strShapefileSlopeClass_LandCover = "SlopeClass_LandCover"
    strTblTPBreakPointsName = "TPBreakPoints"
    strFieldGRIDCODE = "GridCode"
    strFieldAverageExport = "AverageExp"
Else
    'may want to put the following into a text file for the user to load as a configuration
    strTblCatchmentSewageName = "CatchmentSewage"
    strTblCatchPName = "CatchP"
    strTblExportsName = "Exports"
    strTblFlowRouting = "FlowRouting"
    strTblLoadPrecursorName = "LoadPrecursor"
    strTblPerCapitaTPLoadsName = "PerCapitaTPLoads"
    strShapefileSlopeClass_LandCover = "SlopeClass_LandCover"
    strTblTPBreakPointsName = "TPBreakPoints"
    strFieldGRIDCODE = "GridCode"
    strFieldAverageExport = "AverageExp"
    strTblPointSource = "PointSource"
End If

'initialise the Scenario names
strScenarioLocalCatchmentAndNetwork = "LocalCatchment_and_Network_S"
strScenarioTblCatchmentSewageName = "CatchmentSewage_S"
strScenarioTblCatchPName = "CatchP_S"
strScenarioTblExportsName = "Exports_S"
strScenarioTblLoadPrecursorName = "LoadPrecursor_S"
strScenarioTblPerCapitaTPLoadsName = "PerCapitaTPLoads_S"
strScenarioTblTPBreakPointsName = "TPBreakPoints_S"
strScenarioTbl = "Scenario"
strScenarioTblPointSource = "PointSource_S"

strScenarioTableNames(0) = strScenarioTblCatchmentSewageName
strScenarioTableNames(1) = strScenarioTblCatchPName
strScenarioTableNames(2) = strScenarioTblExportsName
strScenarioTableNames(3) = strScenarioTblLoadPrecursorName
strScenarioTableNames(4) = strScenarioTblPerCapitaTPLoadsName
strScenarioTableNames(5) = strScenarioTblTPBreakPointsName
strScenarioTableNames(6) = strScenarioTbl

blnModifySewageLoad = False
blnModifyOtherPointSourceLoad = False
'#######################################################################################
'Initialise by getting the tables, names defined above
'#######################################################################################
Dim blnTable1Found As Boolean
Dim blnTable2Found As Boolean
Dim blnTable3Found As Boolean
Dim blnTable4Found As Boolean
Dim blnTable5Found As Boolean
Dim blnTable6Found As Boolean
Dim blnTable7Found As Boolean
blnTable1Found = False
blnTable2Found = False
blnTable3Found = False
blnTable4Found = False
blnTable5Found = False
blnTable6Found = False
blnTable7Found = False
Dim blnTable1Found_S As Boolean
Dim blnTable2Found_S As Boolean
Dim blnTable3Found_S As Boolean
Dim blnTable4Found_S As Boolean
Dim blnTable5Found_S As Boolean
Dim blnTable6Found_S As Boolean
Dim blnTable7Found_S As Boolean
blnTable1Found_S = False
blnTable2Found_S = False
blnTable3Found_S = False
blnTable4Found_S = False
blnTable5Found_S = False
blnTable6Found_S = False
blnTable7Found_S = False

For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblLoadPrecursorName Then
        Set pLoadPrecursorTable = pTabColl.StandaloneTable(intTempCounter)
        blnTable1Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblCatchPName Then
        Set pCatchPTable = pTabColl.StandaloneTable(intTempCounter)
        blnTable2Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblExportsName Then
        Set pExportsTable = pTabColl.StandaloneTable(intTempCounter)
        blnTable3Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblCatchmentSewageName Then
        Set pCatchmentSewageTable = pTabColl.StandaloneTable(intTempCounter)
        blnTable4Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblPerCapitaTPLoadsName Then
        Set pPerCapitaTPLoads = pTabColl.StandaloneTable(intTempCounter)
        blnTable5Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblTPBreakPointsName Then
        Set pTPBreakPoints = pTabColl.StandaloneTable(intTempCounter)
        blnTable6Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strTblPointSource Then   'added 11.08.2015
        Set pPointSourceTable = pTabColl.StandaloneTable(intTempCounter)
        blnTable7Found = True
    End If
'and the scenario tables
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblLoadPrecursorName Then
        Set pLoadPrecursorTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable1Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblCatchPName Then
        Set pCatchPTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable2Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblExportsName Then
        Set pExportsTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable3Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblCatchmentSewageName Then
        Set pCatchmentSewageTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable4Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblPerCapitaTPLoadsName Then
        Set pPerCapitaTPLoads_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable5Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblTPBreakPointsName Then
        Set pTPBreakpoints_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable6Found_S = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblPointSource Then
        Set pPointSourceTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable7Found_S = True
    End If
Next

If Not blnTable1Found Then
    MsgBox "Table " & strTblLoadPrecursorName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable2Found Then
    MsgBox "Table " & strTblCatchPName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable3Found Then
    MsgBox "Table " & strTblExportsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable4Found Then
    MsgBox "Table " & strTblCatchmentSewageName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable5Found Then
    MsgBox "Table " & strTblPerCapitaTPLoadsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable6Found Then
    MsgBox "Table " & strTblTPBreakPointsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable7Found Then
    MsgBox "Table " & strTblPointSource & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable1Found_S Then
    MsgBox "Table " & strScenarioTblLoadPrecursorName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable2Found_S Then
    MsgBox "Table " & strScenarioTblCatchPName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable3Found_S Then
    MsgBox "Table " & strScenarioTblExportsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable4Found_S Then
    MsgBox "Table " & strScenarioTblCatchmentSewageName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable5Found_S Then
    MsgBox "Table " & strScenarioTblPerCapitaTPLoadsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable6Found_S Then
    MsgBox "Table " & strScenarioTblTPBreakPointsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable7Found_S Then
    MsgBox "Table " & strScenarioTblPointSource & " is not found.", vbCritical
    Exit Sub
End If

'#######################################################################################
'Initialise the catchment
'#######################################################################################
Dim pCompositeLayer As ICompositeLayer

For intTempCounter = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(intTempCounter) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(intTempCounter)
        For j = 0 To pCompositeLayer.Count - 1
        If TypeOf pCompositeLayer.Layer(j) Is IFeatureLayer Then
            If Not pCompositeLayer.Layer(j).Name Like "*RASTER*" And Not pCompositeLayer.Layer(j).Name Like "*STREETVIEW*" Then
                'Debug.Print "1 " & pCompositeLayer.Layer(j).Name
                cboCatchment.AddItem pCompositeLayer.Layer(j).Name
                cboCatchmentScenario.AddItem pCompositeLayer.Layer(j).Name
                Set pFLayerCatchment = pCompositeLayer.Layer(j)
                Set pFClassCatchment = pFLayerCatchment.FeatureClass
                cboCatchment.AddItem pCompositeLayer.Layer(j).Name
                cboCatchment.Text = pCompositeLayer.Layer(j).Name
            End If
        End If
        Next j
    Else ' not grouped layer
        If TypeOf pMap.Layer(intTempCounter) Is IFeatureLayer Then
            If DBExists(pMap.Layer(intTempCounter)) Then
                If ((Not pMap.Layer(intTempCounter).Name Like "*RASTER*") And (Not pMap.Layer(intTempCounter).Name Like "*STREETVIEW*")) Then
                    cboCatchment.AddItem pMap.Layer(intTempCounter).Name
                    'Debug.Print "2 " & pMap.Layer(intTempCounter).Name
                    cboCatchmentScenario.AddItem pMap.Layer(intTempCounter).Name
                        
                    Set pFLayerCatchment = pMap.Layer(intTempCounter)
                    Set pFClassCatchment = pFLayerCatchment.FeatureClass
                    If pFLayerCatchment.Name Like "LocalCatchment*" And pFLayerCatchment.Name Like "*k" Then
                        cboCatchment.Text = pMap.Layer(intTempCounter).Name
                    End If
                    If pFLayerCatchment.Name Like "LocalCatchment*S" Then
                        cboCatchmentScenario.Text = pMap.Layer(intTempCounter).Name
                    End If
                End If
            End If
            If ItExists(pMap.Layer(intTempCounter)) Then
                If Not pMap.Layer(intTempCounter).Name Like "*RASTER*" And Not pMap.Layer(intTempCounter).Name Like "*STREETVIEW*" Then
                cboCatchment.AddItem pMap.Layer(intTempCounter).Name
                'Debug.Print "3 " & pMap.Layer(intTempCounter).Name
                cboCatchmentScenario.AddItem pMap.Layer(intTempCounter).Name
                    Set pFLayerCatchment = pMap.Layer(intTempCounter)
                    Set pFClassCatchment = pFLayerCatchment.FeatureClass
                    If pFLayerCatchment.Name Like "LocalCatchment*" And pFLayerCatchment.Name Like "*k" Then
                        cboCatchment.Text = pMap.Layer(intTempCounter).Name
                    End If
                End If
            End If
        End If
    End If
Next

'initialise the lvwCatchmentInfo
lvwCatchmentInfo.View = lvwReport
lvwCatchmentRelationships1.View = lvwReport
lvwCatchmentRelationships2.View = lvwReport
lstViewSupplement.Arrange = lvwNone
lstViewSupplement.View = lvwReport
Set lvwCatchmentRelationships1.SmallIcons = ImageList_Catchment
Set lvwCatchmentRelationships2.SmallIcons = ImageList_Catchment

'#######################################################################################
'initialise the SEPA monitoring table
'#######################################################################################
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    cboSEPAmonitoring.AddItem pTabColl.StandaloneTable(intTempCounter).Name
    If pTabColl.StandaloneTable(intTempCounter).Name Like "*_detailed_*" Then
        cboSEPAmonitoring.Value = pTabColl.StandaloneTable(intTempCounter).Name
        'cboSEPAmonitoring.Value = "SEPA_detailed_Loch_WB_classification"
    End If
Next

'#######################################################################################
'initialise the GB Lakes Water Body ID LUT table
'#######################################################################################
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    cboGBLakesWBID_LUT.AddItem pTabColl.StandaloneTable(intTempCounter).Name
    If pTabColl.StandaloneTable(intTempCounter).Name Like "*lakes*" Or pTabColl.StandaloneTable(intTempCounter).Name Like "*Lakes*" _
    Or pTabColl.StandaloneTable(intTempCounter).Name Like "*GB*" Then
        cboGBLakesWBID_LUT.Value = pTabColl.StandaloneTable(intTempCounter).Name
    End If
Next

'#######################################################################################
'initialise the SEPA classification concentration statistic table
'#######################################################################################
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    cboClassConcStat.AddItem pTabColl.StandaloneTable(intTempCounter).Name
    If pTabColl.StandaloneTable(intTempCounter).Name Like "*loch_class*" Then
        cboClassConcStat.Value = pTabColl.StandaloneTable(intTempCounter).Name
    End If
Next

'#######################################################################################
'Populate varArrayExportsTable with pExportsTable contents
'0 = MatchCode, 1 = LCOVCODE, 2= SlopeCode, 3 = LCOVDESC, 4 = Min, 5 = Max, 6 = Average
'#######################################################################################
ReDim varArrayExportsTable(90, 7)
intMatchCode = pExportsTable.Table.FindField("MatchCode")
If intMatchCode = -1 Then
    MsgBox "Cannot find field MatchCode in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intLCOVCODE = pExportsTable.Table.FindField("LCOVCODE")
If intLCOVCODE = -1 Then
    MsgBox "Cannot find field LCOVCODE in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intSlopeCode = pExportsTable.Table.FindField("SlopeCode")
If intSlopeCode = -1 Then
    MsgBox "Cannot find field SlopeCode in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intLCOVDESC = pExportsTable.Table.FindField("LCOVDESC")
If intLCOVDESC = -1 Then
    MsgBox "Cannot find field LCOVDESC in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intMin = pExportsTable.Table.FindField("Min_")
If intMin = -1 Then
    MsgBox "Cannot find field Min in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intMax = pExportsTable.Table.FindField("Max_")
If intMax = -1 Then
    MsgBox "Cannot find field Max in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If
intAverage = pExportsTable.Table.FindField("Average")
If intAverage = -1 Then
    MsgBox "Cannot find field Average in the " & strTblExportsName & " table", vbCritical
    Exit Sub
End If

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
Dim pCursor As ICursor
Dim pRow As IRow
Set pCursor = pExportsTable.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
Dim i As Integer
i = 0
While Not pRow Is Nothing
    varArrayExportsTable(i, 0) = pRow.Value(intMatchCode)
    varArrayExportsTable(i, 1) = pRow.Value(intLCOVCODE)
    varArrayExportsTable(i, 2) = pRow.Value(intSlopeCode)
    varArrayExportsTable(i, 3) = pRow.Value(intLCOVDESC)
    varArrayExportsTable(i, 4) = pRow.Value(intMin)
    varArrayExportsTable(i, 5) = pRow.Value(intMax)
    varArrayExportsTable(i, 6) = pRow.Value(intAverage)
    i = i + 1
    Set pRow = pCursor.NextRow
Wend

'populate the combobox that will be used to choose different land cover
For i = 0 To 28
    cboResolveAreaDifference.AddItem varArrayExportsTable(i, 3)
Next

'the start-up splash screen display
Sleep 750

For i = 1 To 22
    Image_Splash.Left = i * 14
    Image_Splash.Top = i * 10
    Image_Splash.Width = 642 - (i * 28)
    Image_Splash.Height = 485 - (i * 22)
    DoEvents
    Sleep 25
Next

intCatchP_PField = pCatchPTable.Table.FindField("P")
If intCatchP_PField = -1 Then
    MsgBox "Cannot find field P in the tblCatchP table", vbCritical
    Exit Sub
End If
intCatchP_GBLAKES_IDField = pCatchPTable.Table.FindField("GBLAKES_ID")
If intCatchP_GBLAKES_IDField = -1 Then
    MsgBox "Cannot find field GBLAKES_ID in the tblCatchP table", vbCritical
    Exit Sub
End If
intCatchP_LCOVDESCField = pCatchPTable.Table.FindField("LCOVDESC")
If intCatchP_LCOVDESCField = -1 Then
    MsgBox "Cannot find field LCOVDESC in the tblCatchP table", vbCritical
    Exit Sub
End If
intCatchP_AreaField = pCatchPTable.Table.FindField("Area")
If intCatchP_AreaField = -1 Then
    MsgBox "Cannot find field Area in the tblCatchP table", vbCritical
    Exit Sub
End If

Image_Splash.Visible = False
Display.Visible = True
MultiPage1.Visible = True
MultiPage1.Value = 0 'force Multipage to open on the first page
frameLoadFromMaster.Visible = True
tglMasterOrScenario.Visible = True
cmdAbout.Visible = True
cmdEnlarge.Visible = True
cmdReduce.Visible = True

InitialiseReport

'populate the additional point source drop down - Fish Farms, Birds, Other point source load
'these are just included to clarify what this is for
cboPointSourceType.Text = "Fish farms"
cboPointSourceType.AddItem "Fish farms"
cboPointSourceType.AddItem "Birds"
cboPointSourceType.AddItem "Other point source"

'######################################################################################
'to run batch script enable the line below
'######################################################################################
'RunBatchProcesses  'this will process everything using the modified batch exporting - so no jpeg or pdf.

End Sub
Private Sub cboCatchment_Click()
'#######################################################################################
'Give the use the option to choose another catchment shapefile
'#######################################################################################
Dim iOIDList() As Long
Dim lonRowsInTable As Long
Dim lonRowScroller As Long
Dim intCounter As Long
Dim blnFID As Boolean
Dim blnText As Boolean
Dim blnUnsuitable As Boolean
Dim blnFound As Boolean
'find the Shapefile selected in the drop down
FindFLayer pFLayerCatchment, cboCatchment.Text, blnFound


If pFLayerCatchment Is Nothing Then
    MsgBox "Please select a suitable catchment layer."
    Exit Sub
End If

Set pFClassCatchment = pFLayerCatchment.FeatureClass
If pFClassCatchment Is Nothing Then
    MsgBox "Please select a layer for input."
    Exit Sub
End If

Set pDisplayTableCatchment = pFLayerCatchment
Set pTableCatchment = pDisplayTableCatchment.DisplayTable
'Debug.Print "pTableCatchment is taken from " & cboCatchment.Text
Set pFieldsCatchment = pTableCatchment.Fields
Dim intLochAreaHaField As Integer
For intCounter = 0 To pFieldsCatchment.FieldCount - 1
    If pFieldsCatchment.Field(intCounter).Name Like "*GBLAKES_ID" Then
        Set pFieldCatchmentGBLAKES_ID = pFieldsCatchment.Field(intCounter)
    End If
    If pFieldsCatchment.Field(intCounter).Name Like "*SiteName" Then
        Set pFieldCatchmentSiteName = pFieldsCatchment.Field(intCounter)
    End If
    If pFieldsCatchment.Field(intCounter).Name Like "*NETWORK" Then
        Set pFieldCatchmentNetwork = pFieldsCatchment.Field(intCounter)
    End If
Next

cmdGetCatchmentInfo.Enabled = False
cboGBLAKES_IDs.Enabled = False
cboWBID.Enabled = False
cmdIntersectUserShapefile.Enabled = False
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label49.Enabled = False
Label50.Enabled = False
Label51.Enabled = False
txtOutputShapefileName.Enabled = False
cboUserShapefile.Enabled = False
cboSlopeLCoverShapefile.Enabled = False
lblScenarioSaveWarning.Visible = True
lblScenarioSaveWarning2.Visible = True
lblReportSaveWarning.Visible = True
Frame1Scenario.Visible = False
Frame2UserLandCover.Visible = False
chkProduceResultsTable.Enabled = False
chkProduceResultsTable.Value = False
chkProduceResultsCSV.Enabled = False
chkProduceResultsCSV.Value = False
End Sub
Function ResidenceTime(A As Double, zbar As Double, r As Double) As Double
'Calculate residence time
    ResidenceTime = (A * zbar) / r
    'Debug.Print "Residence Time calculated as: " & ResidenceTime
End Function
Function ReturnWFD_WB_ID(lonGB_WB_ID As Double) As Double
'Use the look up table to convert between the GB lakes and Water Body ID's
Dim lonCounter As Long
For lonCounter = 0 To UBound(pGBLakes_WBID_Array, 1)
    If pGBLakes_WBID_Array(lonCounter, 1) = lonGB_WB_ID Then
        ReturnWFD_WB_ID = pGBLakes_WBID_Array(lonCounter, 0)
        Exit For
    End If
Next
End Function
Function ReturnGBLakes_ID(lonGBLAKES_ID As Long) As Double
'Use the look up table to convert between the GB lakes and Water Body ID's
Dim lonCounter As Long
For lonCounter = 0 To UBound(pGBLakes_WBID_Array, 1)
    If pGBLakes_WBID_Array(lonCounter, 0) = lonGBLAKES_ID Then
        ReturnGBLakes_ID = pGBLakes_WBID_Array(lonCounter, 1)
        Exit For
    End If
Next
End Function
Function ReturnSEPA_Status(ByVal lonGB_ID As Double) As String
'get the WFD_WB_ID using the GBLakes_WBID LUT
'pSEPAmonitoringArray(X, 0) = WATER_BODY_ID ... etc. WATER_BODY_NAME,CLASSIFICATION_YEAR,STATUS"
'pSEPAmonitoringArray(lonCounter, 3) = pRow.Value(pSEPAmonitoringTable.Table.FindField("CLASSIFICATION_YEAR"))
Dim lonCounter As Long
For lonCounter = 0 To UBound(pSEPAmonitoringArray, 1)
    If pSEPAmonitoringArray(lonCounter, 0) = ReturnWFD_WB_ID(lonGB_ID) Then
        ReturnSEPA_Status = pSEPAmonitoringArray(lonCounter, 4)
        Exit For
    End If
Next
If ReturnSEPA_Status = "" Then
    ReturnSEPA_Status = "No status"
End If
End Function
Sub CalcCatchNetRship()
'#######################################################################################
'1 For each item in lonGBLAKES_IDNetworkMatchArray() find its relationship with lonValueToSearch
'  is it Upstream Downstream OtherBranch?
'2 Check the order in intOrderArray() - if it is the highest order in the network then it is downstream of everything else and processing can stop
'3 SecondFirst find the route to the bottom for my selected lonValueToSearch, storing all DS_GBLAKES_IDs
'4 If a checked catchment flows into one of these DS_GBLAKES_IDs then it is on a separate branch
'#######################################################################################
Dim blnMaxOrderGBLAKES_ID As Boolean
Dim blnChosenGBLAKES_IDFound As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim intOrder As Integer
Dim lonWorkingGBLAKES_ID As Long
Dim blnWhileMaxGBLAKES_IDNotFound As Boolean
Dim intWhileLoopCounter As Integer

blnMaxOrderGBLAKES_ID = False
intOrder = 0
maxOrder = 0

'Sort the GBLAKES_ID array of the network lonGBLAKES_IDNetworkMatchArray(i) using its order array lonOrderMatchArray(i)
DoubleSortArray lonOrderMatchArray, lonGBLAKES_IDNetworkMatchArray

For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray)
    If lonOrderMatchArray(i) > maxOrder Then
        maxOrder = lonOrderMatchArray(i)
        lonGBLAKES_IDWithMaxOrder = lonGBLAKES_IDNetworkMatchArray(i)
    End If
    If lonGBLAKES_IDNetworkMatchArray(i) = lonChosenGBLAKES_ID Then
        intOrder = lonOrderMatchArray(i)
    End If
Next

If intOrder >= maxOrder Then
    blnMaxOrderGBLAKES_ID = True
End If

Erase CatchNetRship()
ReDim CatchNetRship(UBound(lonGBLAKES_IDNetworkMatchArray), 26)

If blnMaxOrderGBLAKES_ID Then
    'the chosen catchment is the highest order so all the other catchments in the network are upstream - easy!!!
    For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray)
        CatchNetRship(i, 0) = lonGBLAKES_IDNetworkMatchArray(i)
        If lonGBLAKES_IDNetworkMatchArray(i) = lonChosenGBLAKES_ID Then
            CatchNetRship(i, 1) = "Match"
        Else
            CatchNetRship(i, 1) = "Is Upstream of"
        End If
    Next
Else
'#######################################################################################
'Find the network relationship with lonChosenGBLAKES_ID for all other GBLAKES_IDs (where the chosen GBLAKES_ID is not maxorder)
'if intOrder is 0 then all others are either downstream or OtherBranch
'step through all catchments in network and find out if they Is Upstream of or Is Downstream From or Separate Branch
'must be two phases - step down from the lonChosenGBLAKES_ID - if find the working GBLAKES_ID then = Is Downstream from
                   '- step down from the working GBLAKES_ID - if find the lonChosenGBLAKES_ID then = Is Upstream of
                   '- if neither then it must be a separate branch
'#######################################################################################

'#######################################################################################
'1. Step down through each site code to see if they "Is Upstream of"
'#######################################################################################
    For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray, 1)
    'read GBLAKES_ID from the previously created list of GBLAKES_IDs
        CatchNetRship(i, 0) = lonGBLAKES_IDNetworkMatchArray(i)
        lonWorkingGBLAKES_ID = lonGBLAKES_IDNetworkMatchArray(i) 'this is the variable we will process
        blnChosenGBLAKES_IDFound = False  'keep searching until find the lonChosenGBLAKES_ID in the right of the tblFlowRouting - arrayFlowRouting(x,1)
        'test the order of the lonWorkingGBLAKES_ID - if it is MaxOrder then it is Downstream
        If lonOrderMatchArray(i) = maxOrder Then
            blnChosenGBLAKES_IDFound = True
            'it is the max order, it must therefore be downstream from everything
            CatchNetRship(i, 1) = "Is Downstream from"
        End If
        While Not blnChosenGBLAKES_IDFound
        'step through arrayFlowRouting(), following any paths to find the relation to lonChosenGBLAKES_ID for each of the catchments in lonGBLAKES_IDNetworkMatchArray()
            For j = 0 To UBound(arrayFlowRouting)
            If arrayFlowRouting(j, 0) = lonWorkingGBLAKES_ID Then
            'have found an upstream match, check whether the DS is lonChosenGBLAKES_ID
                If arrayFlowRouting(j, 1) = lonChosenGBLAKES_ID Then
                    'have found the relationship with lonChosenGBLAKES_ID
                    CatchNetRship(i, 1) = "Is Upstream of"
                    blnChosenGBLAKES_IDFound = True
                    Exit For
                Else
                    'set the lonWorkingGBLAKES_ID to the downstream pair and keep looking
                    lonWorkingGBLAKES_ID = arrayFlowRouting(j, 1)
                    'if the matching arrayFlowRouting(j, 1) is the max order then this will loop infinitely
                    'the search for the lonChosenGBLAKES_ID has reached the end without finding the chosen so stop the while, don't set CatchNetRship()
                    If lonWorkingGBLAKES_ID = lonGBLAKES_IDWithMaxOrder Then
                        blnChosenGBLAKES_IDFound = True
                    End If
                    Exit For
                End If
            End If
            Next
        Wend
    Next

'#######################################################################################
'2. Step down from the lonChosenGBLAKES_ID to set those which it flows to "Is Downstream from", keep going until find lonGBLAKES_IDWithMaxOrder
'#######################################################################################
    blnWhileMaxGBLAKES_IDNotFound = True
    lonWorkingGBLAKES_ID = lonChosenGBLAKES_ID
    intWhileLoopCounter = 0
    While blnWhileMaxGBLAKES_IDNotFound
        'step through arrayFlowRouting(), following any paths to find the relation to lonChosenGBLAKES_ID for each of the catchments in lonGBLAKES_IDNetworkMatchArray()
        For j = 0 To UBound(arrayFlowRouting)
            If arrayFlowRouting(j, 0) = lonWorkingGBLAKES_ID Then
                'find the location in CatchNetRship(i, 0)
                For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray)
                    If lonGBLAKES_IDNetworkMatchArray(i) = lonWorkingGBLAKES_ID And lonGBLAKES_IDNetworkMatchArray(i) <> lonChosenGBLAKES_ID Then
                        If arrayFlowRouting(j, 1) = lonGBLAKES_IDWithMaxOrder Then
                            For k = 0 To UBound(lonGBLAKES_IDNetworkMatchArray, 1)
                                If CatchNetRship(k, 0) = lonGBLAKES_IDWithMaxOrder Then
                                    CatchNetRship(k, 1) = "Is Downstream from"
                                End If
                            Next
                        End If
                        CatchNetRship(i, 1) = "Is Downstream from"
                    End If
                Next
                lonWorkingGBLAKES_ID = arrayFlowRouting(j, 1)
                If lonWorkingGBLAKES_ID = lonGBLAKES_IDWithMaxOrder Then
                    blnWhileMaxGBLAKES_IDNotFound = False
                End If
                Exit For
            End If
        Next
        intWhileLoopCounter = intWhileLoopCounter + 1
        If intWhileLoopCounter = 5000 Then
        'this is here to trap a Catchment that does not flow into another, although this should be trapped elsewhere
            MsgBox "The program appears to be stuck in a loop, contact the providers of this tool if this message recurs.", vbCritical
            Exit Sub
        End If
    Wend

'#######################################################################################
'3. Assign others as "Separate Branch" & tidy up
'#######################################################################################
    For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray, 1)
        If CatchNetRship(i, 1) = "" Then
            CatchNetRship(i, 1) = "Separate Branch"
        End If
        If CatchNetRship(i, 0) = lonChosenGBLAKES_ID Then
            CatchNetRship(i, 1) = "Chosen GBLAKES_ID"
        End If
    Next
End If

DoEvents

'#######################################################################################
'4. Create the output display
'#######################################################################################
PopulateListViewCatchmentRelationships

'the output of this sub is:
'CatchNetRship(i, 0/1) - 0 is the GBLAKES_ID, one record pair for each Catchment in the matching network for the chosen GBLAKES_ID
'                                     - 1 is the relation ship between the chosen GBLAKES_ID and each Catchment - ASCII either
'                                     "Separate Branch", "Chosen GBLAKES_ID", ""Is Downstream from", "Is Upstream of"
End Sub
Sub PopulateListViewCatchmentRelationships()
lvwCatchmentRelationships2.Visible = False
lvwCatchmentRelationships1.Visible = True
lvwCatchmentRelationships1.Height = intTopListviewBoxesHeight
lvwCatchmentRelationships1.Top = intTopListviewBoxesTop

lvwCatchmentRelationships1.ColumnHeaders.Clear
Dim intArrayColumnWidths1(5) As Integer
intArrayColumnWidths1(1) = 45
intArrayColumnWidths1(2) = 108
intArrayColumnWidths1(3) = 97
intArrayColumnWidths1(4) = 60
intArrayColumnWidths1(5) = 64
Dim strArrayColumnHeadings1(5) As String
strArrayColumnHeadings1(1) = "Site"
strArrayColumnHeadings1(2) = "Site Name"
strArrayColumnHeadings1(3) = "Relationship to chosen"
strArrayColumnHeadings1(4) = "Chosen site"
strArrayColumnHeadings1(5) = "Water body ID"

Dim i As Integer
For i = 1 To 5
   Set CatchmentRelateColumnHeaders1 = lvwCatchmentRelationships1.ColumnHeaders.Add()
   CatchmentRelateColumnHeaders1.Text = strArrayColumnHeadings1(i)
   CatchmentRelateColumnHeaders1.Width = intArrayColumnWidths1(i)
Next

lvwCatchmentRelationships1.ListItems.Clear
For i = 0 To UBound(lonGBLAKES_IDNetworkMatchArray)
    Set List_Item1 = lvwCatchmentRelationships1.ListItems.Add
    List_Item1 = CatchNetRship(i, 0)
    If CatchNetRship(i, 1) = "Match" Or CatchNetRship(i, 1) = "Chosen GBLAKES_ID" Then
        List_Item1.SubItems(1) = ReturnSitename(CLng(CatchNetRship(i, 0)))
        List_Item1.SubItems(2) = "Chosen GBLAKES_ID"
        List_Item1.SubItems(3) = ""
        If ReturnWFD_WB_ID(CLng(CatchNetRship(i, 0))) <> 0 Then
            List_Item1.SubItems(4) = ReturnWFD_WB_ID(CLng(CatchNetRship(i, 0)))
        Else
            List_Item1.SubItems(4) = "-"
        End If
    Else
        List_Item1.SubItems(1) = ReturnSitename(CLng(CatchNetRship(i, 0)))
        List_Item1.SubItems(2) = CatchNetRship(i, 1)
        List_Item1.SubItems(3) = lonChosenGBLAKES_ID
        If ReturnWFD_WB_ID(CLng(CatchNetRship(i, 0))) <> 0 Then
            List_Item1.SubItems(4) = ReturnWFD_WB_ID(CLng(CatchNetRship(i, 0)))
        Else
            List_Item1.SubItems(4) = "-"
        End If
    End If
Next
End Sub
Sub InitialiseScenarioSave()
'#######################################################################################
'Set up the scenario
'#######################################################################################
lblScenarioSaveWarning.Visible = False
lblScenarioSaveWarning2.Visible = False
lblReportSaveWarning.Visible = False
Frame1Scenario.Visible = True

'initialise the environment
Set pMxDoc = ThisDocument
Set pMap = pMxDoc.FocusMap
Set pTabColl = pMap

Dim pDataset As IDataset
Dim strPathToSourceDataGDB As String
Dim intTempCounter As Integer

'initialise by getting the tables, names defined in 'Activate'
Dim blnTable1Found As Boolean
Dim blnTable2Found As Boolean
Dim blnTable3Found As Boolean
Dim blnTable4Found As Boolean
Dim blnTable5Found As Boolean
Dim blnTable6Found As Boolean
Dim blnTable7Found As Boolean
Dim blnTable8Found As Boolean
blnTable1Found = False
blnTable2Found = False
blnTable3Found = False
blnTable4Found = False
blnTable5Found = False
blnTable6Found = False
blnTable7Found = False
blnTable8Found = False

'inform the user where the output will go:
Label32.Caption = "The data for your scenario will be saved in the following tables: " & strScenarioTblCatchmentSewageName _
                    & ", " & strScenarioTblCatchPName & ", " & strScenarioTblExportsName & ", " & strScenarioTblLoadPrecursorName _
                    & ", " & strScenarioTblPerCapitaTPLoadsName & ", " & strScenarioTblTPBreakPointsName _
                    & ", " & strScenarioTbl

'track where the tables are coming from - want to make sure only one occurence of each of these tables
'pass through the scenario tables first to find out if there is more than one scenario GDB loaded
Dim i As Integer
Dim j As Integer
j = 0
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    For i = 0 To 6
        If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTableNames(i) Then
            Set pDataset = pTabColl.StandaloneTable(intTempCounter)
            strPathToSourceDataGDB = pDataset.Workspace.PathName
            ReDim Preserve strListofGDBContainingScenarioTables(j)
            strListofGDBContainingScenarioTables(j) = pDataset.Workspace.PathName
            j = j + 1
        End If
    Next
Next
cmdCreateScenario.Enabled = True
For i = 1 To UBound(strListofGDBContainingScenarioTables)
    If strListofGDBContainingScenarioTables(i) <> strListofGDBContainingScenarioTables(i - 1) Then
    'there is more than one GDB loaded with scenario tables present
        MsgBox "Warning. You appear to have more than one scenario GDB loaded. This tool works with only a single scenario GDB." _
        & vbCrLf & "You will not be able to save a scenario.", vbCritical
        Label32.Caption = "Please ensure that only one scenario GDB is loaded. You will need to select 'Get catchment info.' again."
        cmdCreateScenario.Enabled = False
        Exit Sub
    End If
Next
 
Label32.Caption = Label32.Caption & ". " & vbCrLf & vbCrLf & "These tables are in the " & GetCategory(pDataset.Workspace) _
& " which is located at:" & vbCrLf & strListofGDBContainingScenarioTables(0)

For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblCatchmentSewageName Then
        Set pCatchmentSewageTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable1Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblCatchPName Then
        Set pCatchPTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable2Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblExportsName Then
        Set pExportsTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable3Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblLoadPrecursorName Then
        Set pLoadPrecursorTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable4Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblPerCapitaTPLoadsName Then
        Set pPerCapitaTPLoads_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable5Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTbl Then
        Set pScenario = pTabColl.StandaloneTable(intTempCounter)
        blnTable6Found = True
    End If
    If pTabColl.StandaloneTable(intTempCounter).Name = strScenarioTblPointSource Then
        Set pPointSourceTable_S = pTabColl.StandaloneTable(intTempCounter)
        blnTable7Found = True
    End If
Next

If Not blnTable1Found Then
    MsgBox strScenarioTblCatchmentSewageName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable2Found Then
    MsgBox strScenarioTblCatchPName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable3Found Then
    MsgBox strScenarioTblExportsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable4Found Then
    MsgBox strScenarioTblLoadPrecursorName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable5Found Then
    MsgBox strScenarioTblPerCapitaTPLoadsName & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable6Found Then
    MsgBox strScenarioTbl & " is not found.", vbCritical
    Exit Sub
End If
If Not blnTable7Found Then
    MsgBox strScenarioTblPointSource & " is not found.", vbCritical
    Exit Sub
End If

'get the fields in the Scenario table strScenarioTbl - pScenario
intScenarioIDField = pScenario.Table.FindField("ScenarioID")
If intScenarioIDField = -1 Then
    MsgBox "Cannot find field ScenarioID", vbCritical
    Exit Sub
End If
intScenarioNameField = pScenario.Table.FindField("ScenarioName")
If intScenarioNameField = -1 Then
    MsgBox "Cannot find field ScenarioName", vbCritical
    Exit Sub
End If
intScenarioCreatorField = pScenario.Table.FindField("ScenarioCreator")
If intScenarioCreatorField = -1 Then
    MsgBox "Cannot find field ScenarioCreator", vbCritical
    Exit Sub
End If
intScenarioCreationDateField = pScenario.Table.FindField("ScenarioCreationDate")
If intScenarioCreationDateField = -1 Then
    MsgBox "Cannot find field ScenarioCreationDate", vbCritical
    Exit Sub
End If
intScenarioCommentField = pScenario.Table.FindField("Comment")
If intScenarioCommentField = -1 Then
    MsgBox "Cannot find field Comment", vbCritical
    Exit Sub
End If
intScenarioRegionField = pScenario.Table.FindField("Region")
If intScenarioRegionField = -1 Then
    MsgBox "Cannot find field Region", vbCritical
    Exit Sub
End If

'#######################################################################################
'Get the number of the last scenario saved, if there are no records then suggest scenario ID = 1
'#######################################################################################
Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
Dim pCursor As ICursor
Dim pRow As IRow

Set pCursor = pScenario.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
If pScenario.Table.RowCount(pQueryFilt) = 0 Then
'get the last ScenarioID in the Scenario table - if none there then give it number 1!
    Label31.Caption = "There are no records currently in your scenario table"
    txtScenarioID = 1
Else
    While Not pRow Is Nothing
        txtScenarioID.Text = pRow.Value(intScenarioIDField) + 1
        txtScenarioID_2.Text = pRow.Value(intScenarioIDField) + 1
        Set pRow = pCursor.NextRow
    Wend
    
End If

'this section may require to be changed if it does not suit the SEPA system
'get the Windows logged in user name and offer it as the Owner name
txtScenarioOwner = Environ("USERNAME")

'get the data in the correct format for txtScenarioDate
txtScenarioDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)

'comment box
If optOriginalData Then
        txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using original base line data."
    Else
        txtScenarioComment = "Created on " & Environ("COMPUTERNAME") & " using scenario data."
End If

'overall check - if there is anything wrong in the above this will not be reached
'if the boolean is not set to true then the scenario saving will be blocked
blnScenarioCanBeSaved = True

If Not blnScenarioCanBeSaved Then
    MsgBox "Your scenario has NOT been saved. Please check your settings and try again.", vbCritical
    Exit Sub
End If

End Sub
Sub CopyCatchmentPolygonsToScenario()
'#######################################################################################
'Copy the polygons from the input shapefile to the scenario shapefile
'no modification of data here - although a user could safely modify this data and use it in subsequent scenarios.
'LocalCatchment_and_Network_S (a shapefile): GBLAKES_ID, Sitename, Order_, Catch_Net
'it is not possible to do a straight copy of features as we wish the scenario features to have a scenario ID
'#######################################################################################

'find the catchment layer
FindFLayer pFLayerCatchment, cboCatchment.Text, False
Set pFClassCatchment = pFLayerCatchment.FeatureClass

Dim pFeatureselection As IFeatureSelection
Dim pQueryFilter As IQueryFilter
Dim pActiveView As IActiveView
Dim varSplit As Variant

If cboGBLAKES_IDs.Text <> "" Then
    varSplit = Split(cboGBLAKES_IDs.Text, " - ")
    lonChosenGBLAKES_ID = varSplit(0)
    strChosenSitename = varSplit(1)
End If

Set pActiveView = pMap

Set pQueryFilter = New QueryFilter
Set pFeatureselection = pFLayerCatchment


If blnScenarioLoaded Then
    If optOriginalData Then
        pQueryFilter.WhereClause = "NETWORK =" & lonChosenNetwork
    Else    'want to choose the scenario that was loaded - not the network .
        pQueryFilter.WhereClause = "ScenarioID =" & cboScenarioID.Text
    End If
Else
    pQueryFilter.WhereClause = "NETWORK =" & lonChosenNetwork
End If

'find the scenario layer
FindFLayer pFLayerCatchment_Scenario, cboCatchmentScenario.Text, False
Set pFClassCatchment_Scenario = pFLayerCatchment_Scenario.FeatureClass

'get the workspace and start editing
Dim pDataset As IDataset
Set pDataset = pFClassCatchment_Scenario

Dim pWorkspace As IWorkspace
Set pWorkspace = pDataset.Workspace

Dim pWorkspaceEdit As IWorkspaceEdit
Set pWorkspaceEdit = pWorkspace

pWorkspaceEdit.StartEditing True
pWorkspaceEdit.StartEditOperation
     
'open a cursor on the input feature class with the given query filter
Dim pFeatCursor As IFeatureCursor
If blnScenarioLoaded Then
    If optOriginalData Then
        Set pFeatCursor = pFLayerCatchment.Search(pQueryFilter, False)
    Else
        Set pFeatCursor = pFLayerCatchment_Scenario.Search(pQueryFilter, False)
    End If
Else
    Set pFeatCursor = pFLayerCatchment.Search(pQueryFilter, False)  'however, if I am saving a scenario load, then that should be what is accessed.
End If
'loop through the input features in the cursor, and insert
'them into the destination feature class.
Dim pFeat As IFeature
Dim pOldFeat As IFeature
Dim pRow As IRow
Dim pFlds As IFields
Dim lSFld As Long
Dim i As Long

Set pRow = pFeatCursor.NextFeature
Do Until pRow Is Nothing
  Set pOldFeat = pRow
  Set pFeat = pFClassCatchment_Scenario.CreateFeature
  Set pFlds = pFeat.Fields
  For i = 2 To pFlds.FieldCount - 1
    If pFlds.Field(i).Name <> "Shape_Length" And pFlds.Field(i).Name <> "OBJECTID" And pFlds.Field(i).Name <> "Shape_Area" And pFlds.Field(i).Name <> "ScenarioID" Then
      lSFld = pRow.Fields.FindField(pFlds.Field(i).Name)
      pFeat.Value(i) = pRow.Value(lSFld)
    End If
    If pFlds.Field(i).Name = "ScenarioID" Then
        pFeat.Value(i) = txtScenarioID.Text
    End If
    'debugging
    If pFlds.Field(i).Name = "SITECODE" Then
        Debug.Print pFeat.Value(i)
    End If
    
  Next i
  Set pFeat.Shape = pOldFeat.ShapeCopy
  pFeat.Store
  Set pRow = pFeatCursor.NextFeature
Loop

pWorkspaceEdit.StopEditOperation
pWorkspaceEdit.StopEditing True

End Sub
Function ColourToDisplay(lon_GBLAKES_ID As Long, dblP As Double, strToolText As String, dblArrayUpDown() As Double, dblDowngradeMark As Double, dblUpgradeMark As Double, blnUpdateAdditionalResults As Boolean) As String
'#######################################################################################
'Use the GBLAKES_ID and the table pTPBreakPoints to determine the colour of the traffic light cell
'calculate s - the sewage load for each catchment
'Note that J up/downgrade is calculated in CalcTP so that dblArrayUpDown(2) and dblArrayUpDown(4) are left undetermined here
'This also processes the Additional Results tab - the boolean is at the end of the qualifiers list to ensure that only the
'modelled data instances of this function are used for this
'#######################################################################################
'format of dblArrayUpgradeDowngrade is (1) = Cap to TP downgrade, (2) = Cap to J down, (3) = Cap to TP up, (4) = Cap to J up

'GetFieldIndices - this is already in CalcTP

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
pQueryFilt.SubFields = "GBLAKES_ID,Reference_Type,HighGood_P,GoodModerate_P,ModeratePoor_P,PoorBad_P"

Dim pCursor As ICursor
Dim pRow As IRow
If blnDataLoadedFromAScenario Then
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lon_GBLAKES_ID & " and ScenarioID = " & lonSelectedScenario
    Set pCursor = pTPBreakpoints_S.Table.Search(pQueryFilt, False)
    intHighGood_PField = pTPBreakpoints_S.Table.FindField("HighGood_P")
    intGoodModerate_PField = pTPBreakpoints_S.Table.FindField("GoodModerate_P")
    intModeratePoor_PField = pTPBreakpoints_S.Table.FindField("ModeratePoor_P")
    intPoorBad_PField = pTPBreakpoints_S.Table.FindField("PoorBad_P")
    intReference_TypeField = pTPBreakpoints_S.Table.FindField("Reference_Type")
Else
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lon_GBLAKES_ID
    Set pCursor = pTPBreakPoints.Table.Search(pQueryFilt, False)
End If
Set pRow = pCursor.NextRow
If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
    lblLine0.Visible = True
    lblLine1.Visible = True
    lblRedLine.Visible = True
    lblAmberLine.Visible = True
    lblGreenLine.Visible = True
    Label60.Visible = True
    Label61.Visible = True
    Label62.Visible = True
    Label63.Caption = "Calculated status of " & lonChosenGBLAKES_ID & ", " & ReturnSitename(lonChosenGBLAKES_ID)
    If ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID)) <> 0 Then
        Label63.Caption = Label63.Caption & ". WBID: " & ReturnWFD_WB_ID(CLng(lonChosenGBLAKES_ID))
        'Else
    End If
    Label63.Visible = True
    lblIndicator00.Visible = True
    lblLowerStatus00.Visible = True
    lblUpperStatus00.Visible = True
    lblLowerValue00.Visible = True
    lblUpperValue00.Visible = True
    lblP00.Visible = True
End If
ColourToDisplay = "None"
Dim intRowCount As Integer
Dim dblUnits As Double
intRowCount = 0
While Not pRow Is Nothing
    If dblP < pRow.Value(intHighGood_PField) Then
        ColourToDisplay = "High"
        strToolText = "High"
        dblArrayUpDown(1) = pRow.Value(intHighGood_PField) - dblP
        dblArrayUpDown(3) = 9999   'cannot upgrade beyond High
        dblArrayUpDown(4) = 9999   'cannot upgrade beyond High
        dblDowngradeMark = pRow.Value(intHighGood_PField)
        dblUpgradeMark = 9999   'cannot upgrade beyond High
        If blnUpdateAdditionalResults Then
            If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
                lblIndicator00.Left = Label61.Left + 20
                lblP00.Caption = Format(dblP, "#.00")
                lblP00.Left = lblIndicator00.Left + 16
                lblLowerStatus00.Caption = "Moderate/Good"
                lblUpperStatus00.Caption = "Good/High"
                lblLowerValue00.Caption = Format(pRow.Value(intGoodModerate_PField), "#.00")
                lblUpperValue00.Caption = Format(pRow.Value(intHighGood_PField), "#.00")
                lblIndicator00.BackColor = ReturnRAG_Colour(pRow.Value(intHighGood_PField), 0, dblP)
            End If
        End If
    End If
    If dblP >= pRow.Value(intHighGood_PField) And dblP < pRow.Value(intGoodModerate_PField) Then
        ColourToDisplay = "Good"
        strToolText = "Good"
        dblArrayUpDown(1) = pRow.Value(intGoodModerate_PField) - dblP
        dblArrayUpDown(3) = dblP - pRow.Value(intHighGood_PField)
        dblArrayUpDown(4) = 9999
        dblDowngradeMark = pRow.Value(intGoodModerate_PField)
        dblUpgradeMark = pRow.Value(intHighGood_PField)
        If blnUpdateAdditionalResults Then
            If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
                dblUnits = (pRow.Value(intGoodModerate_PField) - pRow.Value(intHighGood_PField)) / (Label61.Left - Label60.Left)
                lblIndicator00.Left = Label60.Left + (Label60.Width / 2) + ((pRow.Value(intGoodModerate_PField) - dblP) / dblUnits)
                lblP00.Caption = Format(dblP, "#.00")
                lblP00.Left = lblIndicator00.Left + 16
                lblLowerStatus00.Caption = "Moderate/Good"
                lblUpperStatus00.Caption = "Good/High"
                lblLowerValue00.Caption = Format(pRow.Value(intGoodModerate_PField), "#.00")
                lblUpperValue00.Caption = Format(pRow.Value(intHighGood_PField), "#.00")
                lblIndicator00.BackColor = ReturnRAG_Colour(pRow.Value(intGoodModerate_PField), pRow.Value(intHighGood_PField), dblP)
            End If
        End If
    End If
    If dblP >= pRow.Value(intGoodModerate_PField) And dblP < pRow.Value(intModeratePoor_PField) Then
        ColourToDisplay = "Moderate"
        strToolText = "Moderate"
        dblArrayUpDown(1) = pRow.Value(intModeratePoor_PField) - dblP
        dblArrayUpDown(3) = dblP - pRow.Value(intGoodModerate_PField)
        dblArrayUpDown(4) = 9999
        dblDowngradeMark = pRow.Value(intModeratePoor_PField)
        dblUpgradeMark = pRow.Value(intGoodModerate_PField)
        If blnUpdateAdditionalResults Then
            If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
                dblUnits = (pRow.Value(intModeratePoor_PField) - pRow.Value(intGoodModerate_PField)) / (Label61.Left - Label60.Left)
                lblIndicator00.Left = Label60.Left + (Label60.Width / 2) + ((pRow.Value(intModeratePoor_PField) - dblP) / dblUnits) - (lblIndicator00.Width / 2)
                lblP00.Caption = Format(dblP, "#.00")
                lblP00.Left = lblIndicator00.Left + 16
                lblLowerStatus00.Caption = "Poor/Moderate"
                lblUpperStatus00.Caption = "Moderate/Good"
                lblLowerValue00.Caption = Format(pRow.Value(intModeratePoor_PField), "#.00")
                lblUpperValue00.Caption = Format(pRow.Value(intGoodModerate_PField), "#.00")
                lblIndicator00.BackColor = ReturnRAG_Colour(pRow.Value(intModeratePoor_PField), pRow.Value(intGoodModerate_PField), dblP)
            End If
        End If
    End If
    If dblP >= pRow.Value(intModeratePoor_PField) And dblP < pRow.Value(intPoorBad_PField) Then
        ColourToDisplay = "Poor"
        strToolText = "Poor"
        dblArrayUpDown(1) = pRow.Value(intPoorBad_PField) - dblP
        dblArrayUpDown(3) = dblP - pRow.Value(intModeratePoor_PField)
        dblArrayUpDown(4) = 9999
        dblDowngradeMark = pRow.Value(intPoorBad_PField)
        dblUpgradeMark = pRow.Value(intModeratePoor_PField)
        If blnUpdateAdditionalResults Then
            If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
                dblUnits = (pRow.Value(intPoorBad_PField) - pRow.Value(intModeratePoor_PField)) / (Label61.Left - Label60.Left)
                lblIndicator00.Left = Label60.Left + (Label60.Width / 2) + ((pRow.Value(intPoorBad_PField) - dblP) / dblUnits)
                lblP00.Caption = Format(dblP, "#.00")
                lblP00.Left = lblIndicator00.Left + 16
                lblLowerStatus00.Caption = "Bad/Poor"
                lblUpperStatus00.Caption = "Poor/Moderate"
                lblLowerValue00.Caption = Format(pRow.Value(intPoorBad_PField), "#.00")
                lblUpperValue00.Caption = Format(pRow.Value(intModeratePoor_PField), "#.00")
                lblIndicator00.BackColor = ReturnRAG_Colour(pRow.Value(intPoorBad_PField), 9999, dblP)
            End If
        End If
    End If
    If dblP >= pRow.Value(intPoorBad_PField) Then
        ColourToDisplay = "Bad"
        strToolText = "Bad"
        dblArrayUpDown(1) = 9999
        dblArrayUpDown(3) = dblP - pRow.Value(intPoorBad_PField)
        dblArrayUpDown(4) = 9999
        dblDowngradeMark = 9999
        dblUpgradeMark = pRow.Value(intPoorBad_PField)
        If blnUpdateAdditionalResults Then
            If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
                lblIndicator00.Left = lblLine0.Left
                lblP00.Caption = Format(dblP, "#.00")
                lblP00.Left = lblIndicator00.Left + 16
                lblLowerStatus00.Caption = "Bad/Poor"
                lblUpperStatus00.Caption = "Poor/Moderate"
                lblLowerValue00.Caption = Format(pRow.Value(intPoorBad_PField), "#.00")
                lblUpperValue00.Caption = Format(pRow.Value(intModeratePoor_PField), "#.00")
                lblIndicator00.BackColor = "255"    'red for these - there isn't a lower class boundary
            End If
        End If
    End If
    strTPBreakPointsRefType = pRow.Value(intReference_TypeField)
    Set pRow = pCursor.NextRow
    intRowCount = intRowCount + 1
Wend
If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
    strToolText = strToolText & " - selected site"
End If

If ColourToDisplay = "None" Then
    strTPBreakPointsRefType = "Not available"
    strToolText = ""
End If

End Function
Function DisplayRAG(lon_GBLAKES_ID As Long, dblP As Double, strToolText As String, dblArrayUpDown() As Double, dblDowngradeMark As Double, dblUpgradeMark As Double) As String
'#######################################################################################
'This is a cut down version of the ColourToDisplay function with labels etc. removed and returning RAG
'Use the GBLAKES_ID and the table pTPBreakPoints to determine the colour of the traffic light cell
'calculate s - the sewage load for each catchment
'Note that J up/downgrade is calculated in CalcTP so that dblArrayUpDown(2) is left and dblArrayUpDown(4) are left undetermined here
'#######################################################################################
'format of dblArrayUpgradeDowngrade is (1) = Cap to TP downgrade, (2) = Cap to J down, (3) = Cap to TP up, (4) = Cap to J up

'GetFieldIndices  - this is already in CalcTP

Dim pQueryFilt As IQueryFilter2
Set pQueryFilt = New QueryFilter
pQueryFilt.SubFields = "GBLAKES_ID,Reference_Type,HighGood_P,GoodModerate_P,ModeratePoor_P,PoorBad_P"

Dim pCursor As ICursor
Dim pRow As IRow
If blnDataLoadedFromAScenario Then
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lon_GBLAKES_ID & " and ScenarioID = " & lonSelectedScenario
    Set pCursor = pTPBreakpoints_S.Table.Search(pQueryFilt, False)
    intHighGood_PField = pTPBreakpoints_S.Table.FindField("HighGood_P")
    intGoodModerate_PField = pTPBreakpoints_S.Table.FindField("GoodModerate_P")
    intModeratePoor_PField = pTPBreakpoints_S.Table.FindField("ModeratePoor_P")
    intPoorBad_PField = pTPBreakpoints_S.Table.FindField("PoorBad_P")
    intReference_TypeField = pTPBreakpoints_S.Table.FindField("Reference_Type")
Else
    pQueryFilt.WhereClause = "GBLAKES_ID = " & lon_GBLAKES_ID
    Set pCursor = pTPBreakPoints.Table.Search(pQueryFilt, False)
End If
Set pRow = pCursor.NextRow
DisplayRAG = "None"
Dim intRowCount As Integer
Dim dblUnits As Double
intRowCount = 0
While Not pRow Is Nothing
    If dblP < pRow.Value(intHighGood_PField) Then
        DisplayRAG = ReturnRAG_Colour(pRow.Value(intHighGood_PField), 0, dblP)
        strToolText = DisplayRAG
    End If
    If dblP >= pRow.Value(intHighGood_PField) And dblP < pRow.Value(intGoodModerate_PField) Then
        DisplayRAG = ReturnRAG_Colour(pRow.Value(intGoodModerate_PField), pRow.Value(intHighGood_PField), dblP)
        strToolText = DisplayRAG
    End If
    If dblP >= pRow.Value(intGoodModerate_PField) And dblP < pRow.Value(intModeratePoor_PField) Then
        DisplayRAG = ReturnRAG_Colour(pRow.Value(intModeratePoor_PField), pRow.Value(intGoodModerate_PField), dblP)
        strToolText = DisplayRAG
    End If
    If dblP >= pRow.Value(intModeratePoor_PField) And dblP < pRow.Value(intPoorBad_PField) Then
        DisplayRAG = ReturnRAG_Colour(pRow.Value(intPoorBad_PField), pRow.Value(intModeratePoor_PField), dblP)
        strToolText = DisplayRAG
    End If
    If dblP >= pRow.Value(intPoorBad_PField) Then
        DisplayRAG = "255"    'red for these - there isn't a lower class boundary
        strToolText = DisplayRAG
    End If
    strTPBreakPointsRefType = pRow.Value(intReference_TypeField)
    Set pRow = pCursor.NextRow
    intRowCount = intRowCount + 1
Wend
If lon_GBLAKES_ID = lonChosenGBLAKES_ID Then
    strToolText = strToolText & " - selected site"
End If

If DisplayRAG = "None" Then
    strTPBreakPointsRefType = "Not available"
    strToolText = ""
End If

End Function
Public Function GetCategory(ByVal pWorkspace As IWorkspace) As String
    Dim sClassID As String
    sClassID = pWorkspace.WorkspaceFactory.GetClassID
    Select Case sClassID
    Case "{DD48C96A-D92A-11D1-AA81-00C04FA33A15}" ' pGDB
        GetCategory = "Personal Geodatabase Database"
    Case "{71FE75F0-EA0C-4406-873E-B7D53748AE7E}" ' fGDB
        GetCategory = "File Geodatabase"
    Case "{D9B4FA40-D6D9-11D1-AA81-00C04FA33A15}" ' GDB
        GetCategory = "SDE Database"
    Case "{A06ADB96-D95C-11D1-AA81-00C04FA33A15}" ' Shape
        GetCategory = "Shapefile Workspace"
    Case "{1D887452-D9F2-11D1-AA81-00C04FA33A15}" ' Coverage
        GetCategory = "ArcInfo Coverage Workspace"
    Case Else
        GetCategory = "Unknown Workspace Category"
    End Select
End Function
Function ItExists(pLayer As ILayer) As Boolean
'#######################################################################################
'Function to check whether a source file that a layer is pointing to actually exists
'#######################################################################################
Dim pDataLayer As IDataLayer
Set pDataLayer = pLayer
ItExists = False
Dim pDatasetName As IDatasetName2
Dim pWSName As IWorkspaceName
Dim sFDS As String
Dim pFCName As IFeatureClassName

Set pDatasetName = pDataLayer.DataSourceName
Set pWSName = pDatasetName.WorkspaceName

If TypeOf pDatasetName Is IFeatureClassName Then
    
    Set pFCName = pDatasetName
    If Not pFCName.FeatureDatasetName Is Nothing Then
        sFDS = pFCName.FeatureDatasetName.Name
    End If
    
End If

If Dir$(pWSName.PathName) Like "*mdb*" Then
Else
    If (GetAttr(pWSName.PathName)) = vbDirectory Then   'only allows directories, so no SDE connections
        If Dir$(pWSName.PathName & "\" & pDatasetName.Name & ".shp") <> "" Then ItExists = True
    End If
End If
End Function
Function TableInDB(pTable As ITable) As Boolean
'#######################################################################################
'Function to check whether a source file for a table that a layer is pointing to actually exists
'#######################################################################################
Dim pDataLayer As IDataLayer
Set pDataLayer = pTable
TableInDB = False
Dim pDatasetName As IDatasetName2
Dim pWSName As IWorkspaceName
Dim sFDS As String
Dim pFCName As IFeatureClassName

Set pDatasetName = pDataLayer.DataSourceName
Set pWSName = pDatasetName.WorkspaceName

If TypeOf pDatasetName Is IFeatureClassName Then
    
    Set pFCName = pDatasetName
    If Not pFCName.FeatureDatasetName Is Nothing Then
        sFDS = pFCName.FeatureDatasetName.Name
    End If
    
End If

If Dir$(pWSName.PathName) <> "" Then TableInDB = True

End Function
Function DBExists(pLayer As ILayer) As Boolean
'#######################################################################################
'Function to check whether a source file in a geodatabase that a layer is pointing to actually exists
'#######################################################################################
Dim pDataLayer As IDataLayer
Set pDataLayer = pLayer
DBExists = False
Dim pDatasetName As IDatasetName2
Dim pWSName As IWorkspaceName
Dim sFDS As String
Dim pFCName As IFeatureClassName

Set pDatasetName = pDataLayer.DataSourceName
Set pWSName = pDatasetName.WorkspaceName

If TypeOf pDatasetName Is IFeatureClassName Then
    
    Set pFCName = pDatasetName
    If Not pFCName.FeatureDatasetName Is Nothing Then
        sFDS = pFCName.FeatureDatasetName.Name
    End If
    
End If

If Dir$(pWSName.PathName) <> "" Then DBExists = True

End Function
Function FindFLayer(pFLayerToFind As IFeatureLayer, strFLayerToFind As String, blnFound As Boolean)
'#######################################################################################
'Function to find a feature layer in the table of contents
'#######################################################################################

Dim intLayerCount As Integer
Dim intGroupLayerCount As Integer
Dim pGroupLayer As IGroupLayer
Dim pCompositeLayer As ICompositeLayer

Set pMxDoc = ThisDocument
Set pMap = pMxDoc.FocusMap
blnFound = False
For intLayerCount = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(intLayerCount) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(intLayerCount)
        For intGroupLayerCount = 0 To pCompositeLayer.Count - 1
            If pCompositeLayer.Layer(intGroupLayerCount).Name = strFLayerToFind Then
                Set pFLayerToFind = pCompositeLayer.Layer(intGroupLayerCount)
                blnFound = True
            End If
        Next intGroupLayerCount
    Else
        If pMap.Layer(intLayerCount).Name = strFLayerToFind Then
            Set pFLayerToFind = pMap.Layer(intLayerCount)
            blnFound = True
        End If
    End If
Next intLayerCount

End Function
Function FindLayer(pLayerToFind As ILayer, strLayerNameToFind As String)
'#######################################################################################
'Function to find a layer in the table of contents
'#######################################################################################

Dim intLayerCount As Integer
Dim intGroupLayerCount As Integer
Dim pGroupLayer As IGroupLayer
Dim pCompositeLayer As ICompositeLayer

Set pMxDoc = ThisDocument
Set pMap = pMxDoc.FocusMap

For intLayerCount = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(intLayerCount) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(intLayerCount)
        For intGroupLayerCount = 0 To pCompositeLayer.Count - 1
            If pCompositeLayer.Layer(intGroupLayerCount).Name = strLayerNameToFind Then
                Set pLayerToFind = pCompositeLayer.Layer(intGroupLayerCount)
            End If
        Next intGroupLayerCount
    Else
        If pMap.Layer(intLayerCount).Name = strLayerNameToFind Then
            Set pLayerToFind = pMap.Layer(intLayerCount)
        End If
    End If
Next intLayerCount

End Function
Private Function DeleteFeaturesInShapefile(strShapeFileName As String)
'#######################################################################################
'Function to delete all features in a selected shapefile
'#######################################################################################

Dim pFLayerInput As IFeatureLayer
Dim pFClassInput As IFeatureClass

Dim intLayerCount As Integer

For intLayerCount = 0 To pMap.LayerCount - 1
    If pMap.Layer(intLayerCount).Name = strShapeFileName Then
        Set pFLayerInput = pMap.Layer(intLayerCount)
        Set pFClassInput = pFLayerInput.FeatureClass
    End If
Next intLayerCount

Dim pQueryFilter As IQueryFilter
Set pQueryFilter = New QueryFilter

Dim pTable As ITable
Set pTable = pFClassInput
On Error Resume Next    'in case there are no searched rows
pTable.DeleteSearchedRows pQueryFilter

End Function
Function LayerLoaded(strLayerName As String, blnLayerLoaded As Boolean, blnRemoveLayer As Boolean)
'#######################################################################################
'Function to confirm that a layer is loaded, with the option to remove it if it is
'#######################################################################################

Dim intLayerCount As Integer
Dim intGroupLayerCount As Integer
Dim pGroupLayer As IGroupLayer
Dim pCompositeLayer As ICompositeLayer

blnLayerLoaded = False
For intLayerCount = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(intLayerCount) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(intLayerCount)
        For intGroupLayerCount = 0 To pCompositeLayer.Count - 1
            If pCompositeLayer.Layer(intGroupLayerCount).Name = strLayerName Then
                blnLayerLoaded = True
                If blnRemoveLayer Then pMap.DeleteLayer pCompositeLayer.Layer(intGroupLayerCount)
            End If
        Next intGroupLayerCount
    Else
        If pMap.Layer(intLayerCount).Name = strLayerName Then
                If blnRemoveLayer Then pMap.DeleteLayer pMap.Layer(intLayerCount)
                blnLayerLoaded = False 'now
                Exit For
        End If
    End If
Next intLayerCount

If blnRemoveLayer And blnLayerLoaded Then pMap.DeleteLayer pMap.Layer(intLayerCount)

End Function
Public Function DeleteField(ByRef pFClass As IFeatureClass, ByVal sFieldName As String) As Boolean
'#######################################################################################
'Function to delete a field in a table
'#######################################################################################

Dim pFields As IFields
Dim pField As IField

DeleteField = False

Dim intFieldID As Integer
intFieldID = pFClass.FindField(sFieldName)
If intFieldID = -1 Then
    MsgBox "Warning the process could not find field " & sFieldName & " to delete.", vbInformation
Else
    Set pFields = pFClass.Fields
    Set pField = pFields.Field(pFClass.FindField(sFieldName))
    
    If Not pField Is Nothing Then
      pFClass.DeleteField pField
      DeleteField = True
    End If
End If

End Function
Private Function GetLCOVDESC(intGridCode As Integer, varArrayInput() As Variant) As String
'#######################################################################################
'Function to look up an array for a land cover description
'#######################################################################################

Dim intCounter As Integer
For intCounter = 0 To UBound(varArrayInput)
    If varArrayInput(intCounter, 0) = intGridCode Then
        GetLCOVDESC = varArrayInput(intCounter, 1)
        Exit For
    End If
Next
End Function
Public Function DoubleSortArray(ByRef TheSortArray As Variant, ByRef TheSecondArray As Variant)
'#######################################################################################
'Function to sort the second and first arrays, using information only in the first
'this is an inverse sort
'#######################################################################################

Dim temp As Variant
Dim temp2 As Variant
Dim X As Integer
Dim Sorted As Boolean
Sorted = False
Do While Not Sorted
    Sorted = True
    For X = 0 To UBound(TheSortArray) - 1
        If TheSortArray(X) < TheSortArray(X + 1) Then
            temp = TheSortArray(X + 1)
            temp2 = TheSecondArray(X + 1)
            TheSortArray(X + 1) = TheSortArray(X)
            TheSecondArray(X + 1) = TheSecondArray(X)
            TheSortArray(X) = temp
            TheSecondArray(X) = temp2
            Sorted = False
        End If
    Next X
Loop
End Function
Function ReturnSitename(ByVal iGBLAKES_ID As Long) As String
'#######################################################################################
'Function to return the sitename for a given GBLAKES_ID
'#######################################################################################

'deal with the Gaelic character set here
'substitute chr(149) for chr(242) etc.
Dim strTemp As String
strTemp = ""
Dim i As Long
For i = 0 To UBound(lonGBLAKES_IDArray)
    If lonGBLAKES_IDArray(i) = iGBLAKES_ID Then
            Dim j As Integer
            For j = 1 To Len(strSitenameArray(i))
                If Asc(Mid(strSitenameArray(i), j, 1)) = 149 Then
                    strTemp = strTemp & Chr(242)
                ElseIf Asc(Mid(strSitenameArray(i), j, 1)) = 130 Then
                     strTemp = strTemp & Chr(233)
                ElseIf Asc(Mid(strSitenameArray(i), j, 1)) = 138 Then
                     strTemp = strTemp & Chr(232)
                ElseIf Asc(Mid(strSitenameArray(i), j, 1)) = 141 Then
                     strTemp = strTemp & Chr(236)
                ElseIf Asc(Mid(strSitenameArray(i), j, 1)) = 151 Then
                     strTemp = strTemp & Chr(250)
                ElseIf Asc(Mid(strSitenameArray(i), j, 1)) = 162 Then
                     strTemp = strTemp & Chr(243)
                Else
                    strTemp = strTemp & Mid(strSitenameArray(i), j, 1)
                End If
            Next
        ReturnSitename = strTemp
        Exit For
    End If
Next
End Function
Function ReturnExportCoeff(ByRef pExportsTab As Variant, ByVal intGridCode As Long) As Double
Dim i As Long
For i = 0 To UBound(pExportsTab, 1)
    If pExportsTab(i, 0) = intGridCode Then
        ReturnExportCoeff = pExportsTab(i, 6)
        Exit For
    End If
Next
End Function
Sub InitialiseReport()
'to capture the location of the .mxd document
Dim j As Integer
Dim strPathToMXD As String
Dim VbProj As VBProject
Dim pApp As IApplication
Set VbProj = Application.Document.VBProject

Dim varArray As Variant
varArray = Split(VbProj.FileName, "\")
Dim strPath As String
strPathToMXD = ""
strPathToMXD = varArray(0)
For j = 1 To UBound(varArray) - 1
    strPathToMXD = strPathToMXD & "\" & varArray(j)
Next
txtOutputReport.Text = strPathToMXD & "\"
txtOutputFile.Text = strPathToMXD & "\"
DoEvents
End Sub
Public Sub ReplaceTitleText(strTextToInsert As String)
'The SEPA title is arial font 22, centre justified, position x = 9.1608, y = 1.6532.
'However the text origin is the lower left corner, so this is re-calculated depending on the width and number of rows.

Dim pMxDocument As IMxDocument
Set pMxDocument = Application.Document

'Switch to Layout Mode
If Not pMxDocument.ActiveView Is pMxDocument.PageLayout Then
    Set pMxDocument.ActiveView = pMxDocument.PageLayout
End If

Dim pLayout As IPageLayout
Set pLayout = pMxDocument.ActiveView
Dim pElement As IElement
Set pGraphicsContainer = pLayout
pGraphicsContainer.Reset
    
Dim pArray As esriSystem.IArray
Set pArray = New esriSystem.Array

Dim pTextElement As ITextElement
Set pElement = pGraphicsContainer.Next
While Not pElement Is Nothing
    If TypeOf pElement Is ITextElement Then
        
        Set pTextElement = pElement
        Dim pElementProps As IElementProperties
        Set pElementProps = pTextElement
        If pElementProps.Type = "Text" Then
            If pTextElement.Symbol.size = "22" And pTextElement.Symbol.Font = "Arial" Then
                pTextElement.Text = strTextToInsert
            End If
            Set pTextElement = Nothing
        End If
    End If
    Set pElement = pGraphicsContainer.Next
Wend

Dim pAV As IActiveView
Set pAV = pMxDocument.PageLayout
pAV.PartialRefresh esriViewGraphics, Nothing, Nothing
    
End Sub
Function NewMapScale(CurMapScale As Long) As Long
'Calculate suitable scales to round up to
Select Case CurMapScale
    Case 0 To 100
        NewMapScale = 100
    Case 101 To 250
        NewMapScale = 250
    Case 251 To 500
        NewMapScale = 500
    Case 501 To 750
        NewMapScale = 750
    Case 751 To 1000
        NewMapScale = 1000
    Case 1001 To 1250
        NewMapScale = 1250
    Case 1251 To 1500
        NewMapScale = 1500
    Case 1501 To 1750
        NewMapScale = 1750
    Case 1751 To 2000
        NewMapScale = 2000
    Case 2001 To 2500
        NewMapScale = 2500
    Case 2501 To 3000
        NewMapScale = 3000
    Case 3001 To 3500
        NewMapScale = 3500
    Case 3501 To 4000
        NewMapScale = 4000
    Case 4001 To 500
        NewMapScale = 4500
    Case 4501 To 5000
        NewMapScale = 5000
    Case 5001 To 5500
        NewMapScale = 5500
    Case 5501 To 6000
        NewMapScale = 6000
    Case 6001 To 6500
        NewMapScale = 6500
    Case 6501 To 7000
        NewMapScale = 7000
    Case 7001 To 7500
        NewMapScale = 7500
    Case 7501 To 8000
        NewMapScale = 8000
    Case 8001 To 8500
        NewMapScale = 8500
    Case 8501 To 9000
        NewMapScale = 9000
    Case 9001 To 9500
        NewMapScale = 9500
    Case 9501 To 10000
        NewMapScale = 10000
    Case 10001 To 12500
        NewMapScale = 12500
    Case 12500 To 15000
        NewMapScale = 15000
    Case 15001 To 17500
        NewMapScale = 17500
    Case 17501 To 20000
        NewMapScale = 20000
    Case 20001 To 25000
        NewMapScale = 25000
    Case 25001 To 30000
        NewMapScale = 30000
    Case 30001 To 35000
        NewMapScale = 35000
    Case 35001 To 40000
        NewMapScale = 40000
    Case 40001 To 45000
        NewMapScale = 45000
    Case 45001 To 50000
        NewMapScale = 50000
    Case 50001 To 55000
        NewMapScale = 55000
    Case 55001 To 60000
        NewMapScale = 60000
    Case 60001 To 65000
        NewMapScale = 65000
    Case 65001 To 70000
        NewMapScale = 70000
    Case 70001 To 75000
        NewMapScale = 75000
    Case 75001 To 80000
        NewMapScale = 80000
    Case 80001 To 85000
        NewMapScale = 85000
    Case 85001 To 90000
        NewMapScale = 90000
    Case 90001 To 95000
        NewMapScale = 95000
    Case 95001 To 100000
        NewMapScale = 100000
    Case 100001 To 110000
        NewMapScale = 110000
    Case 110001 To 120000
        NewMapScale = 120000
    Case 120001 To 130000
        NewMapScale = 130000
    Case 130001 To 140000
        NewMapScale = 140000
    Case 140001 To 150000
        NewMapScale = 150000
    Case 150001 To 160000
        NewMapScale = 160000
    Case 160001 To 170000
        NewMapScale = 170000
    Case 170001 To 180000
        NewMapScale = 180000
    Case 180001 To 190000
        NewMapScale = 190000
    Case 190001 To 200000
        NewMapScale = 200000
    Case 200001 To 250000
        NewMapScale = 250000
    Case 250001 To 300000
        NewMapScale = 300000
    Case 300001 To 350000
        NewMapScale = 350000
    Case 350001 To 400000
        NewMapScale = 400000
    Case 400001 To 450000
        NewMapScale = 450000
    Case 450001 To 500000
        NewMapScale = 500000
    Case 500001 To 550000
        NewMapScale = 550000
    Case 550001 To 600000
        NewMapScale = 600000
    Case 600001 To 650000
        NewMapScale = 650000
    Case 650001 To 700000
        NewMapScale = 700000
    Case 700001 To 750000
        NewMapScale = 750000
    Case 750001 To 800000
        NewMapScale = 800000
    Case 800001 To 850000
        NewMapScale = 850000
    Case 850001 To 900000
        NewMapScale = 900000
    Case 900001 To 950000
        NewMapScale = 950000
    Case 950001 To 1000000
        NewMapScale = 1000000
    Case 1000001 To 1250000
        NewMapScale = 1250000
    Case 1250001 To 1500000
        NewMapScale = 1500000
    Case 1500001 To 1750000
        NewMapScale = 1750000
    Case 1750001 To 2000000
        NewMapScale = 2000000
    Case 2000001 To 2250000
        NewMapScale = 2250000
    Case 2250001 To 2500000
        NewMapScale = 2500000
    Case Is > 2500000
        NewMapScale = 2500000
    Case Else
        NewMapScale = CurMapScale
End Select
End Function
Sub AdjustMapFrame(xMin As Double, yMin As Double, xMax As Double, yMax As Double)
Dim pMxDocument As IMxDocument
Dim pElement As IElement
Dim pEnvelope As IEnvelope
Dim pPoint As IPoint
Dim pEnumElement As IEnumElement

'Update 29.08.2017:
'The app was crashing on the GraphicsContainer.Reset action with a new MXD from SEPA
'This new MXD had a different layout and it appears that the code couldn't find elements
'within the graphics container at the hard coded location of 0.6007, 0.672 so I have changed those to
'xmin and ymin and increased the search tolerance to 0.01

Set pMxDocument = ThisDocument
'force it to LayOut so that graphic elements will be found
Set pMxDocument.ActiveView = pMxDoc.PageLayout
Set pGraphicsContainer = pMxDocument.ActiveView ' Page Layout.

'get the map frame
pGraphicsContainer.Reset
Set pElement = pGraphicsContainer.Next
Do While Not pElement Is Nothing
    If TypeOf pElement Is IMapFrame Then
        Set pEnvelope = New Envelope
        pEnvelope.xMin = xMin
        pEnvelope.xMax = xMax
        pEnvelope.yMin = yMin
        pEnvelope.yMax = yMax
        pElement.Geometry = pEnvelope
        'Exit Do
    End If
    Set pElement = pGraphicsContainer.Next
Loop

'the program also uses this function to collapse the map frame for the tabular output, it sets xMax to 0 for this, so use that as a
'filter to avoid altering the title frame
'adjust the title frame
If xMax <> 0 Then
    Set pPoint = New Point
    pPoint.X = xMin '0.6007
    pPoint.Y = yMin '0.672
    pGraphicsContainer.Reset
    Set pEnumElement = pGraphicsContainer.LocateElements(pPoint, 0.01) '001)
    pEnumElement.Reset
    Set pElement = pEnumElement.Next
    Set pEnvelope = pElement.Geometry.Envelope
    pEnvelope.xMax = xMax
    pElement.Geometry = pEnvelope
End If
End Sub
Sub GenerateTable(intNumRows As Integer, lonStartRecord As Long, intPageNumber As Integer, intTotalPages As Integer)
'Create the first page of the report
Dim dblColumnHeight As Double
Dim intRows As Integer
Dim dblCurrentHeight As Double
Dim dblStartHeight As Double
Dim ColumnHeaderText As String
Dim fntSize As Double
Dim marginSize As Double
Const intColumns = 8
Dim strArrayColumnHeaders(intColumns) As String
strArrayColumnHeaders(0) = "GB Lakes ID"
strArrayColumnHeaders(1) = "WB ID"
strArrayColumnHeaders(2) = "Name"
strArrayColumnHeaders(3) = "Stream order"
strArrayColumnHeaders(4) = "TP (" & Chr(181) & "g/l)"
strArrayColumnHeaders(5) = "Total P load (kg per year)"
strArrayColumnHeaders(6) = "Modelled P status"
strArrayColumnHeaders(7) = "SEPA status" & " (" & pSEPAmonitoringArray(0, 3) & ")"   'return the first record of the table
strArrayColumnHeaders(8) = ""
Dim dblColumnWidth(intColumns) As Double
dblColumnWidth(0) = 1.6272 'site code
dblColumnWidth(1) = 1.7  'WB ID
dblColumnWidth(2) = 6.33 'site name
dblColumnWidth(3) = 1.8   'order
dblColumnWidth(4) = 2.2 'TP
dblColumnWidth(5) = 2.2 'J
dblColumnWidth(6) = 2.1 'Status
dblColumnWidth(7) = 1.8272 'SEPA status
dblColumnWidth(8) = 0

'########################################################################################################
'Need to restrict the output table to no more than 28 data rows, any more than 28 rows will need to go into a second page (and more)
'########################################################################################################

dblColumnHeight = 0.75 'CDbl(columnHeight.Text)
intRows = CInt(intNumRows)

fntSize = 12 '((dblColumnHeight) * 72) / 4
marginSize = 1 'fntSize * 0.35

Dim tableWidth As Double
tableWidth = (intColumns + 1) * 2.19893 'dblColumnWidth, suitable for 9 columns
Set pGraphicsContainer = pMxDoc.PageLayout

Dim xMax As Double
Dim xMin As Double
Dim yMax As Double
Dim yMin As Double

xMin = 0.6005
xMax = tableWidth + xMin
yMax = 29.0795
yMin = yMax - 1.2 'create height of title strip of summary report

CreateTableHeader xMin, xMax, yMin, yMax, fntSize, marginSize, intPageNumber, intTotalPages, "FirstPage"

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) ' + 3.2984
yMax = yMin
yMin = yMax - 1.7 'determine height of column header for the summary report

Dim i As Long
i = 0
Do Until i = intColumns
    ColumnHeaderText = strArrayColumnHeaders(i)
    CreateColumnHeader xMin, xMax, yMin, yMax, ColumnHeaderText, fntSize, marginSize
    xMin = xMax
    i = i + 1
    xMax = xMax + dblColumnWidth(i)
Loop
Dim iRows As Long

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) 'xMin

iRows = lonStartRecord  'this was 0, but to take account of multipage outputs, this is required
i = 0

fntSize = ((dblColumnHeight) * 72) / 0.4
marginSize = 2 'fntSize * 0.25

If intRows > 28 Then
    intRows = iRows + 28
    If intRows > UBound(CatchNetRship, 1) + 1 Then
     intRows = UBound(CatchNetRship, 1) + 1
    End If
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
    If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
           yMin = yMax - (dblColumnHeight * 1.5)
       Else
           yMin = yMax - dblColumnHeight
    End If
        Do Until i = intColumns
            CreateColumn xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
Else
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
        If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
               yMin = yMax - (dblColumnHeight * 1.5)
           Else
               yMin = yMax - dblColumnHeight
        End If
        Do Until i = intColumns
            CreateColumn xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
End If

End Sub
Sub GenerateTableMeas(intNumRows As Integer, lonStartRecord As Long, intPageNumber As Integer, intTotalPages As Integer)
'Create the first page of the report
Dim dblColumnHeight As Double
Dim intRows As Integer
Dim dblCurrentHeight As Double
Dim dblStartHeight As Double
Dim ColumnHeaderText As String
Dim fntSize As Double
Dim marginSize As Double
Const intColumns = 8
Dim strArrayColumnHeaders(intColumns) As String
strArrayColumnHeaders(0) = "GB Lakes ID"
strArrayColumnHeaders(1) = "WB ID"
strArrayColumnHeaders(2) = "Name"
strArrayColumnHeaders(3) = "Stream order"
strArrayColumnHeaders(4) = "SEPA Meas. TP (" & Chr(181) & "g/l)"
strArrayColumnHeaders(5) = "Calc. P load (kg per year)"
strArrayColumnHeaders(6) = "PLUS+ P status"
strArrayColumnHeaders(7) = "SEPA status" & " (" & pSEPAmonitoringArray(0, 3) & ")"   'return the first record of the table
strArrayColumnHeaders(8) = ""
Dim dblColumnWidth(intColumns) As Double
dblColumnWidth(0) = 1.6272 'site code
dblColumnWidth(1) = 1.7  'WB ID
dblColumnWidth(2) = 6.13 'site name
dblColumnWidth(3) = 1.8   'order
dblColumnWidth(4) = 2.2 'TP
dblColumnWidth(5) = 2.2 'J
dblColumnWidth(6) = 2.0272 'Status
dblColumnWidth(7) = 2.1 'SEPA status
dblColumnWidth(8) = 0

'########################################################################################################
'Need to restrict the output table to no more than 28 data rows,any more than 28 rows will need to go into a second page (and more)
'########################################################################################################

dblColumnHeight = 0.75 'CDbl(columnHeight.Text)
intRows = CInt(intNumRows)

fntSize = 12 '((dblColumnHeight) * 72) / 4
marginSize = 1 'fntSize * 0.35

Dim tableWidth As Double
tableWidth = (intColumns + 1) * 2.19893 'dblColumnWidth, suitable for 9 columns
Set pGraphicsContainer = pMxDoc.PageLayout

Dim xMax As Double
Dim xMin As Double
Dim yMax As Double
Dim yMin As Double

xMin = 0.6005
xMax = tableWidth + xMin
yMax = 29.0795
yMin = yMax - 1.2 'create height of title strip of summary report

CreateTableHeader xMin, xMax, yMin, yMax, fntSize, marginSize, intPageNumber, intTotalPages, "MeasStats"

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) ' + 3.2984
yMax = yMin
yMin = yMax - 1.7 'determine height of column header for the summary report

Dim i As Long
i = 0
Do Until i = intColumns
    ColumnHeaderText = strArrayColumnHeaders(i)
    CreateColumnHeader xMin, xMax, yMin, yMax, ColumnHeaderText, fntSize, marginSize
    xMin = xMax
    i = i + 1
    xMax = xMax + dblColumnWidth(i)
Loop
Dim iRows As Long

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) 'xMin

iRows = lonStartRecord  'this was 0, but to take account of multipage outputs, this is required
i = 0

fntSize = ((dblColumnHeight) * 72) / 0.4
marginSize = 2 'fntSize * 0.25

If intRows > 28 Then
    intRows = iRows + 28
    If intRows > UBound(CatchNetRship, 1) + 1 Then
     intRows = UBound(CatchNetRship, 1) + 1
    End If
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
    If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
           yMin = yMax - (dblColumnHeight * 1.5)
       Else
           yMin = yMax - dblColumnHeight
    End If
        Do Until i = intColumns
            CreateColumnMeasStats xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
Else
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
        If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
               yMin = yMax - (dblColumnHeight * 1.5)
           Else
               yMin = yMax - dblColumnHeight
        End If
        Do Until i = intColumns
            CreateColumnMeasStats xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
End If

End Sub
Sub GenerateTableCapacity(intNumRows As Integer, lonStartRecord As Long, intPageNumber As Integer, intTotalPages As Integer, strCapType As String)
'Create the first page of the report
Dim dblColumnHeight As Double
Dim intRows As Integer
Dim dblCurrentHeight As Double
Dim dblStartHeight As Double
Dim ColumnHeaderText As String
Dim fntSize As Double
Dim marginSize As Double
Const intColumns = 8
Dim strArrayColumnHeaders(intColumns) As String
strArrayColumnHeaders(0) = "GB Lakes ID"
strArrayColumnHeaders(1) = "WB ID"
strArrayColumnHeaders(2) = "Name"
If strCapType = "Modelled" Then
    strArrayColumnHeaders(3) = "RAG"
Else
    strArrayColumnHeaders(3) = "RAG" 'may change to status for measured  - may need both...
End If
strArrayColumnHeaders(4) = "Cap. to down grade TP (" & Chr(181) & "g/l)"
strArrayColumnHeaders(5) = "Cap. to down grade J (kg/year)"
strArrayColumnHeaders(6) = "Cap. to upgrade TP (" & Chr(181) & "g/l)"
strArrayColumnHeaders(7) = "Cap. to upgrade J (kg/year)"
strArrayColumnHeaders(8) = ""
Dim dblColumnWidth(intColumns) As Double
dblColumnWidth(0) = 1.6272 'site code
dblColumnWidth(1) = 1.7  'WB ID
dblColumnWidth(2) = 6.23 'site name
dblColumnWidth(3) = 1.5  'RAG
dblColumnWidth(4) = 2.18 'TP
dblColumnWidth(5) = 2.18 'J
dblColumnWidth(6) = 2.18 'TP
dblColumnWidth(7) = 2.1872 'J
dblColumnWidth(8) = 0

'########################################################################################################
'Need to restrict the output table to no more than 28 data rows,any more than 28 rows will need to go into a second page (and more)
'########################################################################################################

dblColumnHeight = 0.75 'CDbl(columnHeight.Text)
intRows = CInt(intNumRows)

fntSize = 12 '((dblColumnHeight) * 72) / 4
marginSize = 1 'fntSize * 0.35

Dim tableWidth As Double
tableWidth = (intColumns + 1) * 2.19893 'dblColumnWidth, suitable for 9 columns
Set pGraphicsContainer = pMxDoc.PageLayout

Dim xMax As Double
Dim xMin As Double
Dim yMax As Double
Dim yMin As Double

xMin = 0.6005
xMax = tableWidth + xMin
yMax = 29.0795
yMin = yMax - 1.2 'create height of title strip of summary report

CreateTableHeader xMin, xMax, yMin, yMax, fntSize, marginSize, intPageNumber, intTotalPages, strCapType

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) ' + 3.2984
yMax = yMin
yMin = yMax - 2.3 'determine height of column header for the summary report (second row)

Dim i As Long
i = 0
Do Until i = intColumns
    ColumnHeaderText = strArrayColumnHeaders(i)
    CreateColumnHeader xMin, xMax, yMin, yMax, ColumnHeaderText, fntSize, marginSize
    xMin = xMax
    i = i + 1
    xMax = xMax + dblColumnWidth(i)
Loop
Dim iRows As Long

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) 'xMin

iRows = lonStartRecord  'this was 0, but to take account of multipage outputs, this is required
i = 0

fntSize = ((dblColumnHeight) * 72) / 0.4
marginSize = 2 'fntSize * 0.25

If intRows > 28 Then
    intRows = iRows + 28
    If intRows > UBound(CatchNetRship, 1) + 1 Then
     intRows = UBound(CatchNetRship, 1) + 1
    End If
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
    If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
           yMin = yMax - (dblColumnHeight * 1.5)
       Else
           yMin = yMax - dblColumnHeight
    End If
        Do Until i = intColumns
            CreateColumnCap xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i, strCapType
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
Else
    Do Until iRows = intRows
        yMax = yMin
 'deal with long names
        If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
               yMin = yMax - (dblColumnHeight * 1.5)
           Else
               yMin = yMax - dblColumnHeight
        End If
        Do Until i = intColumns
            CreateColumnCap xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i, strCapType
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
End If

End Sub
Sub GenerateOutputTable(intNumRows As Integer, lonStartRecord As Long, intPageNumber As Integer, _
                        intTotalPages As Integer, intColumns As Integer, strArrayColumnHeaders() As String, _
                        dblColumnWidth() As Double, arrayData() As Variant, strOutputType As String)
Dim dblColumnHeight As Double
Dim intRows As Integer
Dim dblCurrentHeight As Double
Dim dblStartHeight As Double
Dim ColumnHeaderText As String
Dim fntSize As Double
Dim marginSize As Double

'########################################################################################################
'Need to restrict the output table to no more than 28 data rows, any more than 28 rows will need to go into a second page (and more)
'########################################################################################################
'want to include a percentage of each land cover
'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P

If strOutputType = "LandCover" Then
Dim intJ As Integer
dblSumLandCover = 0
For intJ = 0 To UBound(arrayCatchPforChosenGBLAKES_ID)
    If optReportOnBaseline Then
        dblSumLandCover = dblSumLandCover + arrayCatchPforChosenGBLAKES_ID(intJ, 3)
    Else
        dblSumLandCover = dblSumLandCover + arrayCatchPforChosenGBLAKES_ID(intJ, 5)
    End If
Next
End If

dblColumnHeight = 0.75
intRows = CInt(intNumRows)

fntSize = 12
marginSize = 1

Dim tableWidth As Double
tableWidth = 19.7904

Set pGraphicsContainer = pMxDoc.PageLayout

Dim xMax As Double
Dim xMin As Double
Dim yMax As Double
Dim yMin As Double

xMin = 0.6005
xMax = tableWidth + xMin
yMax = 29.0795
yMin = yMax - 1.2

CreateTableHeader xMin, xMax, yMin, yMax, fntSize, marginSize, intPageNumber, intTotalPages, strOutputType

xMin = 0.6005
xMax = xMin + dblColumnWidth(0) ' + 3.2984
yMax = yMin ' - dblColumnHeight
yMin = yMax - 1.2
If strOutputType = "Summary" Then
    yMin = yMax - 2.1
End If

Dim i As Long
i = 0
Do Until i = intColumns
    ColumnHeaderText = strArrayColumnHeaders(i)
    CreateColumnHeader xMin, xMax, yMin, yMax, ColumnHeaderText, fntSize, marginSize
    xMin = xMax
    i = i + 1
    xMax = xMax + dblColumnWidth(i)
Loop
Dim iRows As Long

xMin = 0.6005
xMax = xMin + dblColumnWidth(0)
yMax = yMin + dblColumnHeight
yMin = yMax - dblColumnHeight

iRows = lonStartRecord  'this was 0, but to take account of multipage outputs, this is required
i = 0

fntSize = ((dblColumnHeight) * 72) / 0.4
marginSize = 2
If intRows > 28 Then
    intRows = iRows + 28
    If intRows > UBound(arrayData, 1) + 1 Then
     intRows = UBound(arrayData, 1)
    End If
    Do Until iRows = intRows And iRows < (UBound(arrayData, 1) + 1)
        yMax = yMin
 'deal with long land cover types
        If Len(ReturnSitename(CLng(CatchNetRship(iRows, 0)))) > 25 Then
               yMin = yMax - (dblColumnHeight * 1.5)
           Else
               yMin = yMax - dblColumnHeight
        End If
        Do Until i = intColumns
            CreateLCoverPointSummaryColumn xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i, arrayData(), strOutputType
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i) 'dblColumnWidth
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
Else
    Do Until iRows = intRows
    If iRows > UBound(arrayData, 1) Then
        Exit Do
    End If
        yMax = yMin
 'deal with long land cover types
        If strOutputType = "LandCover" Then
            If Len(arrayCatchPforChosenGBLAKES_ID_With_Summary(iRows, 2)) > 25 Then 'modified with arrayCatchPforChosenGBLAKES_ID_With_Summary
               yMin = yMax - (dblColumnHeight * 1.5)
            Else
                yMin = yMax - dblColumnHeight
            End If
        Else
            yMin = yMax - dblColumnHeight
        End If
        Do Until i = intColumns
            CreateLCoverPointSummaryColumn xMin, xMax, yMin, yMax, fntSize, marginSize, iRows, i, arrayData(), strOutputType
            xMin = xMax
            i = i + 1
            xMax = xMax + dblColumnWidth(i)
        Loop
        xMin = 0.6005
        xMax = xMin + dblColumnWidth(0)
        i = 0
        iRows = iRows + 1
    Loop
End If

End Sub
Private Sub CreateColumn(xMin As Double, xMax As Double, yMin As Double, yMax As Double, fntSize As Double, _
                                marginSize As Double, ByVal intIndexToProcess As Integer, ByVal intColumnToProcess As Integer)
 Dim pParaTextElement As ITextElement
 Dim pElement As IElement
 Dim pEnvelope As IEnvelope
 Dim pActiveView As IActiveView
 
 Dim pRGBcolor As IRgbColor
 Dim pTextSymbol As ITextSymbol
 Dim fnt As IFontDisp
 
 Set pEnvelope = New Envelope
 Set pRGBcolor = New RgbColor
 pRGBcolor.Blue = 0
 pRGBcolor.Red = 0
 pRGBcolor.Green = 0
 
 Dim pFontDisp As IFontDisp
 Set pFontDisp = New stdole.StdFont
 pFontDisp.Name = "Arial"
 pFontDisp.Bold = False
 pFontDisp.Underline = False
 Set pTextSymbol = New TextSymbol
 pTextSymbol.Font = pFontDisp
 pTextSymbol.Color = pRGBcolor
 pTextSymbol.size = "12"
 pTextSymbol.HorizontalAlignment = esriTHACenter
 pTextSymbol.VerticalAlignment = esriTVACenter
 
 pEnvelope.xMin = xMin
 pEnvelope.yMin = yMin
 pEnvelope.xMax = xMax
 pEnvelope.yMax = yMax
 
 Set pElement = New ParagraphTextElement
 pElement.Geometry = pEnvelope
 
 Set pParaTextElement = pElement
 pParaTextElement.Symbol = pTextSymbol
 
'headings & CatchNetRship
'strArrayColumnHeaders(0) = "Site"
'strArrayColumnHeaders(1) = "WB ID"
'strArrayColumnHeaders(2) = "Site name"
'strArrayColumnHeaders(3) = "Order"
'strArrayColumnHeaders(4) = "TP (" & Chr(181) & "g/l)"
'strArrayColumnHeaders(5) = "J (kg)"
'strArrayColumnHeaders(6) = "Status"
'strArrayColumnHeaders(7) = "SEPA status"
'strArrayColumnHeaders(8) = ""
Select Case intColumnToProcess
    Case Is = 0
        intColumnToProcess = 0
    Case Is = 1
        intColumnToProcess = 99
    Case Is = 2
        intColumnToProcess = 1
    Case Is = 3
        intColumnToProcess = 2
    Case Is = 4
        intColumnToProcess = 14
    Case Is = 5
        intColumnToProcess = 10
    Case Is = 6
        intColumnToProcess = 21
End Select
If intColumnToProcess = 1 Then
    pParaTextElement.Text = ReturnSitename(CLng(CatchNetRship(intIndexToProcess, 0)))
ElseIf intColumnToProcess = 10 Or intColumnToProcess = 14 Then
    If CatchNetRship(intIndexToProcess, intColumnToProcess) < 0 Then
        pParaTextElement.Text = "0" & Format(CatchNetRship(intIndexToProcess, intColumnToProcess), "#0.0")
    Else
        pParaTextElement.Text = Format(CatchNetRship(intIndexToProcess, intColumnToProcess), "#0.0")
    End If
ElseIf intColumnToProcess = 2 Then
    pParaTextElement.Text = lonOrderMatchArray(intIndexToProcess)
ElseIf intColumnToProcess = 99 Then 'get the WBID
     pParaTextElement.Text = ReturnWFD_WB_ID(CLng(CatchNetRship(intIndexToProcess, 0)))
     If pParaTextElement.Text = 0 Then
        pParaTextElement.Text = "N/A"
     End If
ElseIf intColumnToProcess = 7 Then 'get the SEPA status
    pParaTextElement.Text = ReturnSEPA_Status(CatchNetRship(intIndexToProcess, 0))
    If pParaTextElement.Text = "No status" Then
        pParaTextElement.Text = "N/A"
    End If
Else
    If CatchNetRship(intIndexToProcess, intColumnToProcess) Like "*selected site" Then
        pParaTextElement.Text = Left(CatchNetRship(intIndexToProcess, intColumnToProcess), Len(CatchNetRship(intIndexToProcess, intColumnToProcess)) - 16)
    Else
        pParaTextElement.Text = CatchNetRship(intIndexToProcess, intColumnToProcess)
    End If
End If

Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSymBorder = New SymbolBorder
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(255, 255, 255)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground

pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer

End Sub
Private Sub CreateColumnMeasStats(xMin As Double, xMax As Double, yMin As Double, yMax As Double, fntSize As Double, _
                                marginSize As Double, ByVal intIndexToProcess As Integer, ByVal intColumnToProcess As Integer)
 Dim pParaTextElement As ITextElement
 Dim pElement As IElement
 Dim pEnvelope As IEnvelope
 Dim pActiveView As IActiveView
 
 Dim pRGBcolor As IRgbColor
 Dim pTextSymbol As ITextSymbol
 Dim fnt As IFontDisp
 
 Set pEnvelope = New Envelope
 Set pRGBcolor = New RgbColor
 pRGBcolor.Blue = 0
 pRGBcolor.Red = 0
 pRGBcolor.Green = 0
 
 Dim pFontDisp As IFontDisp
 Set pFontDisp = New stdole.StdFont
 pFontDisp.Name = "Arial"
 pFontDisp.Bold = False
 pFontDisp.Underline = False
 Set pTextSymbol = New TextSymbol
 pTextSymbol.Font = pFontDisp
 pTextSymbol.Color = pRGBcolor
 pTextSymbol.size = "12"
 pTextSymbol.HorizontalAlignment = esriTHACenter
 pTextSymbol.VerticalAlignment = esriTVACenter
 
 pEnvelope.xMin = xMin
 pEnvelope.yMin = yMin
 pEnvelope.xMax = xMax
 pEnvelope.yMax = yMax
 
 Set pElement = New ParagraphTextElement
 pElement.Geometry = pEnvelope
 
 Set pParaTextElement = pElement
 pParaTextElement.Symbol = pTextSymbol
 
'headings & CatchNetRship
'strArrayColumnHeaders(0) = "Site"
'strArrayColumnHeaders(1) = "WB ID"
'strArrayColumnHeaders(2) = "Site name"
'strArrayColumnHeaders(3) = "Order"
'strArrayColumnHeaders(4) = "TP (" & Chr(181) & "g/l)" - measured
'strArrayColumnHeaders(5) = "J (kg)" - back calculated from TP
'strArrayColumnHeaders(6) = "Status" - plus
'strArrayColumnHeaders(7) = "SEPA status"
'strArrayColumnHeaders(8) = ""
Select Case intColumnToProcess
    Case Is = 0
        intColumnToProcess = 0
    Case Is = 1
        intColumnToProcess = 99
    Case Is = 2
        intColumnToProcess = 1
    Case Is = 3
        intColumnToProcess = 2
    Case Is = 4
        intColumnToProcess = 25
    Case Is = 5 'calculate here
        intColumnToProcess = 26
    Case Is = 6
        intColumnToProcess = 21
End Select
If intColumnToProcess = 1 Then
    pParaTextElement.Text = ReturnSitename(CLng(CatchNetRship(intIndexToProcess, 0)))
ElseIf intColumnToProcess = 25 Or intColumnToProcess = 26 Then
    If CatchNetRship(intIndexToProcess, intColumnToProcess) < 1 Then     'CatchNetRship(i, 26)= Back calculated J (total P load) for the SEPA concentration
        pParaTextElement.Text = "0" & Format(CatchNetRship(intIndexToProcess, intColumnToProcess), "#0.0")
    Else
        pParaTextElement.Text = Format(CatchNetRship(intIndexToProcess, intColumnToProcess), "#0.0")
    End If
    If CatchNetRship(intIndexToProcess, intColumnToProcess) = 0 Then
        pParaTextElement.Text = "N/A"
    End If
ElseIf intColumnToProcess = 2 Then
    pParaTextElement.Text = lonOrderMatchArray(intIndexToProcess)
ElseIf intColumnToProcess = 99 Then 'get the WBID
     pParaTextElement.Text = ReturnWFD_WB_ID(CLng(CatchNetRship(intIndexToProcess, 0)))
     If pParaTextElement.Text = 0 Then
        pParaTextElement.Text = "N/A"
     End If
ElseIf intColumnToProcess = 7 Then 'get the SEPA status
    pParaTextElement.Text = ReturnSEPA_Status(CatchNetRship(intIndexToProcess, 0))
    If pParaTextElement.Text = "No status" Then
        pParaTextElement.Text = "N/A"
    End If
     Select Case pParaTextElement.Text
        Case "H"
            pParaTextElement.Text = "High"
        Case "G"
            pParaTextElement.Text = "Good"
        Case "M"
            pParaTextElement.Text = "Moderate"
        Case "P"
            pParaTextElement.Text = "Poor"
        Case "B"
            pParaTextElement.Text = "Bad"
     End Select
Else
    If CatchNetRship(intIndexToProcess, intColumnToProcess) Like "*selected site" Then
        pParaTextElement.Text = Left(CatchNetRship(intIndexToProcess, intColumnToProcess), Len(CatchNetRship(intIndexToProcess, intColumnToProcess)) - 16)
    Else
        pParaTextElement.Text = CatchNetRship(intIndexToProcess, intColumnToProcess)
    End If
End If

Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSymBorder = New SymbolBorder
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(255, 255, 255)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground

pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer

End Sub
Private Sub CreateLCoverPointSummaryColumn(xMin As Double, xMax As Double, yMin As Double, yMax As Double, fntSize As Double, _
                               marginSize As Double, ByVal intIndexToProcess As Integer, ByVal intColumnToProcess As Integer, _
                               arrayData() As Variant, strOutputType As String)
                               
Dim pParaTextElement As ITextElement
Dim pElement As IElement
Dim pEnvelope As IEnvelope
Dim pActiveView As IActiveView

Dim pRGBcolor As IRgbColor
Dim pTextSymbol As ITextSymbol
Dim fnt As IFontDisp

Set pEnvelope = New Envelope
Set pRGBcolor = New RgbColor
pRGBcolor.Blue = 0
pRGBcolor.Red = 0
pRGBcolor.Green = 0

Dim pFontDisp As IFontDisp
Set pFontDisp = New stdole.StdFont
pFontDisp.Name = "Arial"
pFontDisp.Bold = False
pFontDisp.Underline = False
Set pTextSymbol = New TextSymbol
pTextSymbol.Font = pFontDisp
pTextSymbol.Color = pRGBcolor
pTextSymbol.size = "12"
pTextSymbol.HorizontalAlignment = esriTHACenter
pTextSymbol.VerticalAlignment = esriTVACenter

pEnvelope.xMin = xMin
pEnvelope.yMin = yMin
pEnvelope.xMax = xMax
pEnvelope.yMax = yMax

Set pElement = New ParagraphTextElement
pElement.Geometry = pEnvelope

Set pParaTextElement = pElement
pParaTextElement.Symbol = pTextSymbol

'arrayCatchPforChosenGBLAKES_ID(i,j) 0 = GBLAKES_ID, 1 = lcovdesc, 2 = P, 3 = area, 4 = kg/ha, 5 = revised area, 6 = revised kg/ha, 7 = revised P
If strOutputType = "LandCover" Then
    Select Case intColumnToProcess
       Case Is = 100
           pParaTextElement.Text = CLng(arrayData(intIndexToProcess, 0))
       Case Is = 101
           pParaTextElement.Text = ReturnWFD_WB_ID(CLng(arrayData(intIndexToProcess, 0)))
           If pParaTextElement.Text = 0 Then
               pParaTextElement.Text = "-"
           End If
       Case Is = 0
       'process P
           intColumnToProcess = 1
           If optReportOnBaseline Then
               pParaTextElement.Text = arrayData(intIndexToProcess, 2)
           Else
               pParaTextElement.Text = arrayData(intIndexToProcess, 7)
           End If
           If CDbl(pParaTextElement.Text) < 1 Then
               pParaTextElement.Text = "0" & Format(CDbl(pParaTextElement.Text), "#.0")
           Else
               pParaTextElement.Text = Format(CDbl(pParaTextElement.Text), "#.0")
           End If
       Case Is = 1
       'process land cover
           intColumnToProcess = 2
           pParaTextElement.Text = arrayData(intIndexToProcess, 1)
       Case Is = 2
       'process area
           intColumnToProcess = 3
           If optReportOnBaseline Then
               pParaTextElement.Text = Format(CDbl(arrayData(intIndexToProcess, 3)), "#")
           Else
               pParaTextElement.Text = Format(CDbl(arrayData(intIndexToProcess, 5)), "#")
           End If
       Case Is = 3
       'process area percentage
           intColumnToProcess = 3
           If optReportOnBaseline Then
               pParaTextElement.Text = Format(CDbl(arrayData(intIndexToProcess, 3) / dblSumLandCover) * 100, "0.0")
           Else
               pParaTextElement.Text = Format(CDbl(arrayData(intIndexToProcess, 5) / dblSumLandCover) * 100, "0.0")
           End If
       Case Is = 4
       'process user modified boolean
           intColumnToProcess = 3
           If optReportOnScenario Then
               If (arrayData(intIndexToProcess, 7) <> arrayData(intIndexToProcess, 2)) _
               Or (arrayData(intIndexToProcess, 5) <> arrayData(intIndexToProcess, 3)) Then
                   pParaTextElement.Text = "Yes"
               Else
                    If arrayData(intIndexToProcess, 1) = "Total for all land covers" Then
                        If arrayData(intIndexToProcess, 5) <> arrayData(intIndexToProcess, 3) Then
                            pParaTextElement.Text = "Yes"
                        Else
                            pParaTextElement.Text = "N/A"
                        End If
                    Else
                      pParaTextElement.Text = "N/A"
                    End If
               End If
           Else
               pParaTextElement.Text = "N/A"
           End If
    End Select
End If
If strOutputType = "LandCoverSummary" Then

End If
Dim lonIndexChosenGBLAKES_IDSummary As Long
Dim i As Long
lonIndexChosenGBLAKES_IDSummary = 0
If strOutputType = "Summary" Then
'get the index of lonchosenGBLAKES_ID in varCatchmentSewage
    For i = 0 To UBound(arrayData, 1)
        If arrayData(i, 0) = lonChosenGBLAKES_ID Then
            lonIndexChosenGBLAKES_IDSummary = i
        End If
    Next
    Select Case intColumnToProcess
       Case Is = 100 'GBLAKES_ID
           pParaTextElement.Text = CLng(arrayData(intIndexSelectedGBLAKES_ID, 0))
       Case Is = 101 'WB ID
           pParaTextElement.Text = ReturnWFD_WB_ID(CLng(arrayData(intIndexSelectedGBLAKES_ID, 0)))
           If pParaTextElement.Text = 0 Then
               pParaTextElement.Text = "-" 'N/A
           End If
       Case Is = 0 'Sum P
       'these values are the same for either scenario or baseline
            pParaTextElement.Text = Format(dblJSelectedCatchment, "#.0")
       Case Is = 1 'Sum land cover input
       'these values are the same for either scenario or baseline
            pParaTextElement.Text = Format(dblSumLocalInputs, "#.0")
       Case Is = 2 'Sum upstream input
       'these values are the same for either scenario or baseline
            pParaTextElement.Text = Format(dblSumUpstream, "#.0")
       Case Is = 3 'Urban pop - CatchNetRship() - has active data, varCatchmentSewage contains read-in only
           intColumnToProcess = 1
           If optReportOnBaseline Then
               pParaTextElement.Text = varCatchmentSewage(lonIndexChosenGBLAKES_IDSummary, 3)
           Else
               pParaTextElement.Text = arrayData(intIndexSelectedGBLAKES_ID, 18)
           End If
       Case Is = 4 'Urban P
           If optReportOnBaseline Then
               pParaTextElement.Text = varCatchmentSewage(lonIndexChosenGBLAKES_IDSummary, 1)
           Else
               pParaTextElement.Text = arrayData(intIndexSelectedGBLAKES_ID, 16)
           End If
           If CDbl(pParaTextElement.Text) < 1 Then
               pParaTextElement.Text = "0" & Format(CDbl(pParaTextElement.Text), "#.0")
           Else
               pParaTextElement.Text = Format(CDbl(pParaTextElement.Text), "#.0")
           End If
       Case Is = 5 'Rural pop
           If optReportOnBaseline Then
               pParaTextElement.Text = varCatchmentSewage(lonIndexChosenGBLAKES_IDSummary, 4)
           Else
               pParaTextElement.Text = arrayData(intIndexSelectedGBLAKES_ID, 19)
           End If
       Case Is = 6 'Rural P
           If optReportOnBaseline Then
               pParaTextElement.Text = varCatchmentSewage(lonIndexChosenGBLAKES_IDSummary, 2)
           Else
               pParaTextElement.Text = arrayData(intIndexSelectedGBLAKES_ID, 17)
           End If
           If CDbl(pParaTextElement.Text) < 1 Then
               pParaTextElement.Text = "0" & Format(CDbl(pParaTextElement.Text), "#.0")
           Else
               pParaTextElement.Text = Format(CDbl(pParaTextElement.Text), "#.0")
           End If
        Case Is = 7 ' point sources
            pParaTextElement.Text = arrayData(intIndexSelectedGBLAKES_ID, 22)
    End Select
End If

If strOutputType = "PointSources" Then
'get the index of lonchosenGBLAKES_ID in varPointSource
    Select Case intColumnToProcess
       Case Is = 0 'GBLAKES_ID
           pParaTextElement.Text = CLng(arrayData(intIndexToProcess, 0))
       Case Is = 1 'WB_ID
            pParaTextElement.Text = ReturnWFD_WB_ID(CLng(arrayData(intIndexToProcess, 0)))
            If pParaTextElement.Text = 0 Then
               pParaTextElement.Text = "N/A"
            End If
       Case Is = 2 'point source type
           pParaTextElement.Text = arrayData(intIndexToProcess, 1)
       Case Is = 3 'point source amount
           pParaTextElement.Text = arrayData(intIndexToProcess, 2)
    End Select
    If pParaTextElement.Text = "" Or pParaTextElement.Text = "0" Then
        pParaTextElement.Text = " "
    End If
End If

Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSymBorder = New SymbolBorder
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(255, 255, 255)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground

pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer

End Sub
Private Sub CreateColumnCap(xMin As Double, xMax As Double, yMin As Double, yMax As Double, fntSize As Double, _
                                marginSize As Double, ByVal intIndexToProcess As Integer, ByVal intColumnToProcess As Integer, ByVal strCapType As String)
Dim pParaTextElement As ITextElement
Dim pElement As IElement
Dim pEnvelope As IEnvelope
Dim pActiveView As IActiveView

Dim pRGBcolor As IRgbColor
Dim pTextSymbol As ITextSymbol
Dim fnt As IFontDisp

Set pEnvelope = New Envelope
Set pRGBcolor = New RgbColor
pRGBcolor.Blue = 0
pRGBcolor.Red = 0
pRGBcolor.Green = 0

Dim pFontDisp As IFontDisp
Set pFontDisp = New stdole.StdFont
pFontDisp.Name = "Arial"
pFontDisp.Bold = False
pFontDisp.Underline = False
Set pTextSymbol = New TextSymbol
pTextSymbol.Font = pFontDisp
pTextSymbol.Color = pRGBcolor
pTextSymbol.size = "12"
pTextSymbol.HorizontalAlignment = esriTHACenter
pTextSymbol.VerticalAlignment = esriTVACenter

pEnvelope.xMin = xMin
pEnvelope.yMin = yMin
pEnvelope.xMax = xMax
pEnvelope.yMax = yMax

Set pElement = New ParagraphTextElement
pElement.Geometry = pEnvelope

Set pParaTextElement = pElement
pParaTextElement.Symbol = pTextSymbol
 
'headings & CatchNetRship
'strArrayColumnHeaders(0) = "Site"
'strArrayColumnHeaders(1) = "WB ID"
'strArrayColumnHeaders(2) = "Site name"
'strArrayColumnHeaders(3) = "RAG" - or "Meas'd status"
'strArrayColumnHeaders(4) = "Cap to down TP (" & Chr(181) & "g/l)" - or Meas'd
'strArrayColumnHeaders(5) = "Cap to down J (kg)" - or Meas'd
'strArrayColumnHeaders(6) = "Cap to up TP (" & Chr(181) & "g/l)" - or Meas'd
'strArrayColumnHeaders(7) = "Cap to up J (kg)" - or Meas'd
'strArrayColumnHeaders(8) = ""
Dim dblDowngradeMark As Double
Dim dblUpgradeMark As Double
Dim dblTestUpDown As Double
Dim dblArrayUpDown(4) As Double
Dim blnHasMeasuredP As Boolean
blnHasMeasuredP = False
Dim i As Integer
For i = 0 To 4
    dblArrayUpDown(i) = 0
    dblArrayUpDown(i) = 0
Next
Dim dblReturnedWFD_WB_ID As Double
dblReturnedWFD_WB_ID = ReturnWFD_WB_ID(CLng(CatchNetRship(intIndexToProcess, 0)))
Dim dblSEPA_meas_conc As Double
Dim strStatusColour As String
dblSEPA_meas_conc = 0
'ColourToDisplay populates dblArrayUpDown()
If strCapType = "Modelled" Then
    strStatusColour = ColourToDisplay(CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(CatchNetRship(intIndexToProcess, 14)), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark, False)
    DisplayRAG CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(CatchNetRship(intIndexToProcess, 14)), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark
Else 'it's measured
    'find the measured value
    If intColumnToProcess > 2 Then
        dblSEPA_meas_conc = CatchNetRship(intIndexToProcess, 25)
        If dblSEPA_meas_conc <> 0 Then
            blnHasMeasuredP = True
        End If
        strStatusColour = ColourToDisplay(CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(dblSEPA_meas_conc), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark, False)
        DisplayRAG CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(dblSEPA_meas_conc), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark
    End If
End If
If intColumnToProcess = 0 Then
    pParaTextElement.Text = CatchNetRship(intIndexToProcess, intColumnToProcess)
ElseIf intColumnToProcess = 1 Then 'get the WBID
     pParaTextElement.Text = dblReturnedWFD_WB_ID 'ReturnWFD_WB_ID(CLng(CatchNetRship(intIndexToProcess, 0)))
     If pParaTextElement.Text = 0 Then
        pParaTextElement.Text = "N/A"
     End If
ElseIf intColumnToProcess = 2 Then 'get the Site name
    pParaTextElement.Text = ReturnSitename(CLng(CatchNetRship(intIndexToProcess, 0)))
ElseIf intColumnToProcess = 3 Then
    If strCapType = "Modelled" Then 'get the RAG
        Select Case DisplayRAG(CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(CatchNetRship(intIndexToProcess, 14)), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark)
            Case "255"
            pParaTextElement.Text = "Red"
            Case "33023"
            pParaTextElement.Text = "Amber"
            Case "65280"
            pParaTextElement.Text = "Green"
        End Select
        If strStatusColour = "Bad" Then
            pParaTextElement.Text = "Green"
        End If
    Else    'it's measured, so get the measured P status (high, good etc.)
        If blnHasMeasuredP Then
            Select Case DisplayRAG(CLng(CatchNetRship(intIndexToProcess, 0)), CDbl(dblSEPA_meas_conc), "Any", dblArrayUpDown, dblDowngradeMark, dblUpgradeMark)
                Case "255"
                pParaTextElement.Text = "Red"
                Case "33023"
                pParaTextElement.Text = "Amber"
                Case "65280"
                pParaTextElement.Text = "Green"
            End Select
            If strStatusColour = "Bad" Then
                pParaTextElement.Text = "Green"
            End If
        Else
            pParaTextElement.Text = "N/A"
        End If
    End If
ElseIf intColumnToProcess = 4 Then 'Cap to down TP (" & Chr(181) & "g/l)" - or Meas'd
    If strCapType = "Modelled" Then
        If dblArrayUpDown(1) <> 9999 Then
            pParaTextElement.Text = Format(dblArrayUpDown(1), "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
     Else
        If blnHasMeasuredP And dblArrayUpDown(1) <> 9999 Then
            pParaTextElement.Text = Format(dblArrayUpDown(1), "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
     End If
ElseIf intColumnToProcess = 5 Then 'get the "Cap to down J (kg)" - or Meas'd
    If strCapType = "Modelled" Then
        If dblDowngradeMark <> 9999 Then
                'J of break point = sum curr_and_upstr_runoff*(1 + sqr(Tw)))*((downgrd_brk_pt/OECD-a) ^ 1/OECD-b
                'calculate the J (kg) corresponding to the breakpoint
                dblTestUpDown = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblDowngradeMark / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
                'and remove the current TP to calculate the capacity
                dblTestUpDown = dblTestUpDown - CatchNetRship(intIndexToProcess, 10)
                pParaTextElement.Text = Format(dblTestUpDown, "0.0") 'Cap. to down, J
            Else
                pParaTextElement.Text = "N/A"
        End If
    Else
        If blnHasMeasuredP And dblDowngradeMark <> 9999 Then
            dblJ_for_meas_TP = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblSEPA_meas_conc / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
            dblTestUpDown = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblDowngradeMark / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
            dblTestUpDown = dblTestUpDown - dblJ_for_meas_TP
            pParaTextElement.Text = Format(dblTestUpDown, "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
    End If
ElseIf intColumnToProcess = 6 Then 'get the "Cap to up TP (" & Chr(181) & "g/l)" - or Meas'd
    If strCapType = "Modelled" Then
        If dblArrayUpDown(3) <> -9999 And dblArrayUpDown(3) <> 9999 Then
            pParaTextElement.Text = Format(dblArrayUpDown(3) * -1, "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
    Else
        If blnHasMeasuredP And dblArrayUpDown(3) <> 9999 Then
            pParaTextElement.Text = Format(dblArrayUpDown(3) * -1, "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
    End If
ElseIf intColumnToProcess = 7 Then 'get the "Cap to up J (kg)" - or Meas'd
    If strCapType = "Modelled" Then
        If dblUpgradeMark <> -9999 And dblUpgradeMark <> 9999 Then
            dblTestUpDown = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblUpgradeMark / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
            'and remove the current TP
            dblTestUpDown = dblTestUpDown - CatchNetRship(intIndexToProcess, 10) 'CatchNetRship(j, 10) is J (kg), sum of the current and immediate upstream P (total P in the loch in kg)
            pParaTextElement.Text = Format(dblTestUpDown, "0.0") 'Cap. to down, J
        Else
            pParaTextElement.Text = "N/A"
        End If 'j
    Else
        If blnHasMeasuredP And dblArrayUpDown(3) <> 9999 Then
            dblJ_for_meas_TP = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblSEPA_meas_conc / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
            dblTestUpDown = (CatchNetRship(intIndexToProcess, 3) * (1 + Sqr(CatchNetRship(intIndexToProcess, 6)))) * ((dblUpgradeMark / CatchNetRship(intIndexToProcess, 7)) ^ (1 / CatchNetRship(intIndexToProcess, 8)) / 1000000)
            dblTestUpDown = dblTestUpDown - dblJ_for_meas_TP
            pParaTextElement.Text = Format(dblTestUpDown, "0.0")
        Else
            pParaTextElement.Text = "N/A"
        End If
    End If
Else
    If CatchNetRship(intIndexToProcess, intColumnToProcess) Like "*selected site" Then
        pParaTextElement.Text = Left(CatchNetRship(intIndexToProcess, intColumnToProcess), Len(CatchNetRship(intIndexToProcess, intColumnToProcess)) - 16)
    Else
        pParaTextElement.Text = CatchNetRship(intIndexToProcess, intColumnToProcess)
    End If
End If

Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSymBorder = New SymbolBorder
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(255, 255, 255)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground

pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer

End Sub
Private Sub CreateColumnHeader(xMin As Double, xMax As Double, yMin As Double, yMax As Double, columnText As String, fntSize As Double, _
                               marginSize As Double)
Dim pParaTextElement As ITextElement
Dim pElement As IElement
Dim pEnvelope As IEnvelope
Dim pActiveView As IActiveView

Dim pRGBcolor As IRgbColor
Dim pTextSymbol As ITextSymbol
Dim fnt As IFontDisp

Set pEnvelope = New Envelope
Set pRGBcolor = New RgbColor
pRGBcolor.Blue = 0
pRGBcolor.Red = 0
pRGBcolor.Green = 0

Dim pFontDisp As IFontDisp
Set pFontDisp = New stdole.StdFont
pFontDisp.Name = "Arial"
pFontDisp.Bold = True
pFontDisp.Underline = False
Set pTextSymbol = New TextSymbol
pTextSymbol.Font = pFontDisp
pTextSymbol.Color = pRGBcolor
pTextSymbol.size = fntSize
pTextSymbol.HorizontalAlignment = esriTHACenter
pTextSymbol.VerticalAlignment = esriTVACenter

pEnvelope.xMin = xMin
pEnvelope.yMin = yMin
pEnvelope.xMax = xMax
pEnvelope.yMax = yMax

Set pElement = New ParagraphTextElement
pElement.Geometry = pEnvelope

Set pParaTextElement = pElement
pParaTextElement.Symbol = pTextSymbol
pParaTextElement.Text = columnText
Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSymBorder = New SymbolBorder
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(178, 178, 178)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground
pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer
End Sub
Private Sub CreateTableHeader(xMin As Double, xMax As Double, yMin As Double, yMax As Double, fntSize As Double, marginSize As Double, _
                               intPageNum As Integer, intTotalPages As Integer, strReportType As String)
Dim pParaTextElement As ITextElement
Dim pElement As IElement
Dim pEnvelope As IEnvelope
Dim pActiveView As IActiveView

Dim pRGBcolor As IRgbColor
Dim pTextSymbol As ITextSymbol
Dim fnt As IFontDisp

Set pEnvelope = New Envelope

Set pRGBcolor = New RgbColor
pRGBcolor.Blue = 0
pRGBcolor.Red = 0
pRGBcolor.Green = 0

Dim pFontDisp As IFontDisp
Set pFontDisp = New stdole.StdFont
pFontDisp.Name = "Arial"
pFontDisp.Bold = True
pFontDisp.Underline = False
Set pTextSymbol = New TextSymbol
pTextSymbol.Font = pFontDisp
pTextSymbol.Color = pRGBcolor
pTextSymbol.size = fntSize
pTextSymbol.HorizontalAlignment = esriTHACenter
pTextSymbol.VerticalAlignment = esriTVACenter

pEnvelope.xMin = xMin
pEnvelope.yMin = yMin
pEnvelope.xMax = xMax
pEnvelope.yMax = yMax

Set pElement = New ParagraphTextElement
pElement.Geometry = pEnvelope

Set pParaTextElement = pElement
pParaTextElement.Symbol = pTextSymbol

If intTotalPages > 1 Then
    If blnScenarioLoaded Then
        If strReportType = "MeasStats" Then
            pParaTextElement.Text = "Measured and Computed Statistics from PLUS+" & ", page " & intPageNum & " of " & intTotalPages & " for " & ReturnSitename(lonChosenGBLAKES_ID) & " in scenario " & lonSelectedScenario
        Else
            pParaTextElement.Text = "Modelled Statistics from PLUS+" & ", page " & intPageNum & " of " & intTotalPages & " for " & ReturnSitename(lonChosenGBLAKES_ID) & " in scenario " & lonSelectedScenario
        End If
    Else
        If strReportType = "MeasStats" Then
            pParaTextElement.Text = "Measured and Computed Statistics from PLUS+" & ", page " & intPageNum & " of " & intTotalPages & " for " & ReturnSitename(lonChosenGBLAKES_ID)
        Else
            pParaTextElement.Text = "Modelled Statistics from PLUS+" & ", page " & intPageNum & " of " & intTotalPages & " for " & ReturnSitename(lonChosenGBLAKES_ID)
        End If
    End If
Else
    If blnScenarioLoaded Then
        If strReportType = "MeasStats" Then
            pParaTextElement.Text = "Measured and Computed Statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments for Scenario " & lonSelectedScenario
        Else
            pParaTextElement.Text = "Modelled Statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments for Scenario " & lonSelectedScenario
        End If
    Else
        If strReportType = "MeasStats" Then
            pParaTextElement.Text = "Measured and Computed Statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments"
        Else
            pParaTextElement.Text = "Modelled Statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments"
        End If
    End If
End If

If strReportType = "LandCover" Then
    If blnScenarioLoaded Then
        pParaTextElement.Text = "Land cover statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " in Scenario " & lonSelectedScenario
    Else
        pParaTextElement.Text = "Land cover statistics from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID)
    End If
End If

If strReportType = "Summary" Then
    If blnScenarioLoaded Then
        pParaTextElement.Text = "Summary of all P loads for " & ReturnSitename(lonChosenGBLAKES_ID) & " in Scenario " & lonSelectedScenario
    Else
        pParaTextElement.Text = "Summary of all P loads for " & ReturnSitename(lonChosenGBLAKES_ID)
    End If
End If

If strReportType = "Modelled" Then
    If blnScenarioLoaded Then
        pParaTextElement.Text = "Modelled capacity to up/down grade for " & ReturnSitename(lonChosenGBLAKES_ID) & " in Scenario " & lonSelectedScenario & " and connected catchments in Scenario " & lonSelectedScenario
    Else
        pParaTextElement.Text = "Modelled capacity to up/down grade for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments"
    End If
End If

If strReportType = "Measured" Then
    If blnScenarioLoaded Then
        pParaTextElement.Text = "Measured capacity to up/down grade for " & ReturnSitename(lonChosenGBLAKES_ID) & " in Scenario " & lonSelectedScenario & " and connected catchments in Scenario " & lonSelectedScenario
    Else
        pParaTextElement.Text = "Measured capacity to up/down grade for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments"
    End If
End If

If strReportType = "PointSources" Then
    If blnScenarioLoaded Then
        pParaTextElement.Text = "Point source loads from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments in Scenario " & lonSelectedScenario
    Else
        pParaTextElement.Text = "Point source loads from PLUS+ for " & ReturnSitename(lonChosenGBLAKES_ID) & " and connected catchments"
    End If
End If

Dim pFrameProp As IFrameProperties
Dim pBorder As IBorder
Dim pBackGround As ISymbolBackground
Dim pSymBorder As ISymbolBorder
Dim pSymBackground As ISymbolBackground
Dim pLineColor As IRgbColor
Dim pLineSymbol As ILineSymbol
Dim pSFS As ISimpleFillSymbol
Dim pColor As IRgbColor

Set pFrameProp = pElement
Set pLineColor = New RgbColor
Set pLineSymbol = New SimpleLineSymbol
Set pSFS = New SimpleFillSymbol
Set pColor = New RgbColor
Set pSymBackground = New SymbolBackground

pLineColor.RGB = RGB(0, 0, 0)
pLineSymbol.Color = pLineColor
pLineSymbol.Width = 0.2
pSFS.Style = esriSFSSolid
pColor.RGB = RGB(200, 200, 200)
pSFS.Color = pColor
pSFS.Outline = pLineSymbol
pSymBackground.FillSymbol = pSFS

Set pBackGround = pSymBackground

pFrameProp.Background = pBackGround

Dim pMarginProp As IMarginProperties
Dim pColumnProp As IColumnProperties
Set pColumnProp = pElement
Set pMarginProp = pElement
pMarginProp.Margin = marginSize
pColumnProp.Count = 1
pColumnProp.Gap = 0

pGraphicsContainer.AddElement pElement, 0
Set pActiveView = pGraphicsContainer

End Sub
Function Translation(strInput As String) As String
'#######################################################################################
'Function to return the sitename for a given GBLAKES_ID
'#######################################################################################

'deal with the Gaelic character set here
'substitute chr(149) for chr(242)
Dim strTemp As String
strTemp = ""
Dim j As Integer
For j = 1 To Len(strInput)
    If Asc(Mid(strInput, j, 1)) = 149 Then
        strTemp = strTemp & Chr(242)
    ElseIf Asc(Mid(strInput, j, 1)) = 130 Then
         strTemp = strTemp & Chr(233)
    ElseIf Asc(Mid(strInput, j, 1)) = 138 Then
         strTemp = strTemp & Chr(232)
    ElseIf Asc(Mid(strInput, j, 1)) = 141 Then
         strTemp = strTemp & Chr(236)
    ElseIf Asc(Mid(strInput, j, 1)) = 151 Then
         strTemp = strTemp & Chr(250)
    ElseIf Asc(Mid(strInput, j, 1)) = 162 Then
         strTemp = strTemp & Chr(243)
    Else
        strTemp = strTemp & Mid(strInput, j, 1)
    End If
Next
Translation = strTemp
End Function
Private Sub SetOutputQuality(pActiveView As IActiveView, iResampleRatio As Long)
Dim pMap As IMap
Dim pGraphicsContainer As IGraphicsContainer
Dim pElement As IElement
Dim pOutputRasterSettings As IOutputRasterSettings
Dim pMapFrame As IMapFrame
Dim pTmpActiveView As IActiveView

If TypeOf pActiveView Is IMap Then
  Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
  pOutputRasterSettings.ResampleRatio = iResampleRatio
ElseIf TypeOf pActiveView Is IPageLayout Then
  
  'assign ResampleRatio for PageLayout
  Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
  pOutputRasterSettings.ResampleRatio = iResampleRatio
  
  'and assign ResampleRatio to the Maps in the PageLayout
  Set pGraphicsContainer = pActiveView
  pGraphicsContainer.Reset
  Set pElement = pGraphicsContainer.Next
  Do While Not pElement Is Nothing
    If TypeOf pElement Is IMapFrame Then
      Set pMapFrame = pElement
      Set pTmpActiveView = pMapFrame.Map
      Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
      pOutputRasterSettings.ResampleRatio = iResampleRatio
    End If
    DoEvents
    Set pElement = pGraphicsContainer.Next
  Loop
  Set pMap = Nothing
  Set pMapFrame = Nothing
  Set pGraphicsContainer = Nothing
  Set pTmpActiveView = Nothing
End If
Set pOutputRasterSettings = Nothing

End Sub
Sub SwitchReportToBaseline()
optReportOnBaseline.Enabled = True
optReportOnBaseline.Value = True
optReportOnScenario.Enabled = False
optReportOnScenario.Value = False
End Sub
Sub SwitchReportToScenario()
optReportOnBaseline.Enabled = False
optReportOnBaseline.Value = False
optReportOnScenario.Enabled = True
optReportOnScenario.Value = True
End Sub
Sub LoadSepaMonitoringIntoArray()
'#######################################################################################
'get the SEPA classification data and read into an array
'this contains a list of water bodies and their observed status.
'#######################################################################################
Dim intTempCounter As Integer
Dim pQueryFilt As IQueryFilter2
Dim pCursor As ICursor
Dim pRow As IRow
Dim lonCounter As Long
Dim intField As Integer
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = cboSEPAmonitoring Then
        Set pSEPAmonitoringTable = pTabColl.StandaloneTable(intTempCounter)
    End If
Next

'check the fields and exit with an error if any are missing
intField = -1
intField = pSEPAmonitoringTable.Table.FindField("WATER_BODY_ID")
If intField = -1 Then
    MsgBox "Warning, could not find the field WATER_BODY_ID in the table " & cboSEPAmonitoring & ". Exiting.", vbCritical
    Exit Sub
End If
'WATER_BODY_NAME may not be present in all versions of table, so removing
'intField = -1
'intField = pSEPAmonitoringTable.Table.FindField("WATER_BODY_NAME")
'If intField = -1 Then
'    MsgBox "Warning, could not find the field WATER_BODY_NAME in the table " & cboSEPAmonitoring & ". Exiting.", vbCritical
'    Exit Sub
'End If
intField = -1
intField = pSEPAmonitoringTable.Table.FindField("CLASSIFICATION_YEAR")
If intField = -1 Then
    MsgBox "Warning, could not find the field CLASSIFICATION_YEAR in the table " & cboSEPAmonitoring & ". Exiting.", vbCritical
    Exit Sub
End If
intField = -1
intField = pSEPAmonitoringTable.Table.FindField("STATUS")
If intField = -1 Then
    MsgBox "Warning, could not find the field STATUS in the table " & cboSEPAmonitoring & ". Exiting.", vbCritical
    Exit Sub
End If

Set pQueryFilt = New QueryFilter
'resize the pSEPAmonitoringArray to match the size of pSEPAmonitoringTable
ReDim pSEPAmonitoringArray(pSEPAmonitoringTable.Table.RowCount(pQueryFilt), 4)
pQueryFilt.SubFields = "WATER_BODY_ID,CLASSIFICATION_YEAR,STATUS"
Set pCursor = pSEPAmonitoringTable.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
lonCounter = 0
While Not pRow Is Nothing
    pSEPAmonitoringArray(lonCounter, 0) = pRow.Value(pSEPAmonitoringTable.Table.FindField("WATER_BODY_ID"))
    pSEPAmonitoringArray(lonCounter, 1) = 0 'pRow.Value(pSEPAmonitoringTable.Table.FindField("WATER_BODY_NAME"))
    pSEPAmonitoringArray(lonCounter, 2) = 0 'pRow.Value(pSEPAmonitoringTable.Table.FindField("SUB_BASIN_DISTRICT"))
    pSEPAmonitoringArray(lonCounter, 3) = pRow.Value(pSEPAmonitoringTable.Table.FindField("CLASSIFICATION_YEAR"))
    pSEPAmonitoringArray(lonCounter, 4) = pRow.Value(pSEPAmonitoringTable.Table.FindField("STATUS"))
    Set pRow = pCursor.NextRow
    lonCounter = lonCounter + 1
Wend
End Sub
Sub LoadSepaClassConcStat()
Dim intTempCounter As Integer
Dim pQueryFilt As IQueryFilter2
Dim pCursor As ICursor
Dim pRow As IRow
Dim lonCounter As Long
Dim intField As Integer
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = cboClassConcStat Then
        Set pSEPAClassConcStatTable = pTabColl.StandaloneTable(intTempCounter)
    End If
Next

If (pSEPAClassConcStatTable Is Nothing) Then
    MsgBox cboClassConcStat.Text & "not found.", vbCritical
End If

'pSEPAClassConcStatArray()
'check the fields and exit with an error if any are missing
'•   "Year" - currently 2009
'•   "Point classification result" - actual concentration data in micrograms / L
'•   "Water Body ID"
'•   "Class_ID" - this is the current WFD class (1 = high, 5 = bad)

intField = -1
intField = pSEPAClassConcStatTable.Table.FindField("YEAR_")
If intField = -1 Then
    MsgBox "Warning, could not find the field YEAR_ in the table " & cboClassConcStat & ". Exiting.", vbCritical
    Exit Sub
End If
intField = -1
intField = pSEPAClassConcStatTable.Table.FindField("POINT_CLASSIFICATION_RESULT")
If intField = -1 Then
    MsgBox "Warning, could not find the field POINT_CLASSIFICATION_RESULT in the table " & cboClassConcStat & ". Exiting.", vbCritical
    Exit Sub
End If
intField = -1
intField = pSEPAClassConcStatTable.Table.FindField("WATER_BODY_ID")
If intField = -1 Then
    MsgBox "Warning, could not find the field WATER_BODY_ID in the table " & cboClassConcStat & ". Exiting.", vbCritical
    Exit Sub
End If
intField = -1
intField = pSEPAClassConcStatTable.Table.FindField("CLASS_ID")
If intField = -1 Then
    MsgBox "Warning, could not find the field CLASS_ID in the table " & cboClassConcStat & ". Exiting.", vbCritical
    Exit Sub
End If
Set pQueryFilt = New QueryFilter
'resize the pSEPAClassConcStatArray to match the size of pSEPAmonitoringTable
ReDim pSEPAClassConcStatArray(pSEPAClassConcStatTable.Table.RowCount(pQueryFilt), 3)
pQueryFilt.SubFields = "YEAR_,WATER_BODY_ID,POINT_CLASSIFICATION_RESULT,CLASS_ID"
Set pCursor = pSEPAClassConcStatTable.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
lonCounter = 0
While Not pRow Is Nothing
    pSEPAClassConcStatArray(lonCounter, 0) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("WATER_BODY_ID"))
    pSEPAClassConcStatArray(lonCounter, 1) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("YEAR_"))
    pSEPAClassConcStatArray(lonCounter, 2) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("POINT_CLASSIFICATION_RESULT"))
    pSEPAClassConcStatArray(lonCounter, 3) = pRow.Value(pSEPAClassConcStatTable.Table.FindField("CLASS_ID"))
    Set pRow = pCursor.NextRow
    lonCounter = lonCounter + 1
Wend


End Sub
Sub CreateGBLakesTable()
'#######################################################################################
'get the GB Lakes Water Body ID LUT and read into an array
'#######################################################################################
Dim intTempCounter As Integer
Dim pQueryFilt As IQueryFilter2
Dim pCursor As ICursor
Dim pRow As IRow
Dim lonCounter As Long
Dim intField As Integer
For intTempCounter = 0 To pTabColl.StandaloneTableCount - 1
    If pTabColl.StandaloneTable(intTempCounter).Name = cboGBLakesWBID_LUT Then
        Set pGBLakes_WBID_Table = pTabColl.StandaloneTable(intTempCounter)
    End If
Next

'check the fields and exit with an error if any are missing
intField = -1
intField = pGBLakes_WBID_Table.Table.FindField("WFD_WB_ID")
If intField = -1 Then
    MsgBox "Warning, could not find the field WFD_WB_ID in the table " & cboGBLakesWBID_LUT & ". Exiting.", vbCritical
    Exit Sub
End If
intField = -1
intField = pGBLakes_WBID_Table.Table.FindField("GB_WB_ID")
If intField = -1 Then
    MsgBox "Warning, could not find the field GB_WB_ID in the table " & cboGBLakesWBID_LUT & ". Exiting.", vbCritical
    Exit Sub
End If

Set pQueryFilt = New QueryFilter
'resize the pGBLakes_WBID_Array to match the size of pGBLakes_WBID_Table
ReDim pGBLakes_WBID_Array(pGBLakes_WBID_Table.Table.RowCount(pQueryFilt), 1)
pQueryFilt.SubFields = "WFD_WB_ID,GB_WB_ID" 'change to WFD_WB_ID, GB_WB_ID
Set pCursor = pGBLakes_WBID_Table.Table.Search(pQueryFilt, False)
Set pRow = pCursor.NextRow
lonCounter = 0
While Not pRow Is Nothing
    pGBLakes_WBID_Array(lonCounter, 0) = pRow.Value(pGBLakes_WBID_Table.Table.FindField("WFD_WB_ID")) 'change to WFD_WB_ID
    pGBLakes_WBID_Array(lonCounter, 1) = pRow.Value(pGBLakes_WBID_Table.Table.FindField("GB_WB_ID"))
    Set pRow = pCursor.NextRow
    lonCounter = lonCounter + 1
Wend

End Sub
Function ReturnRAG_Colour(ByVal dblLowerClassBoundary As Double, ByVal dblUpperClassBoundary As Double, ByVal dbl_dblP As Double) As String
'#######################################################################################
'Function to return the Red Amber Green Colour
'#######################################################################################

Dim strRAG As String
strRAG = ""
If dbl_dblP > dblLowerClassBoundary - ((dblLowerClassBoundary - dblUpperClassBoundary) * 0.03) Then
    strRAG = "255"
ElseIf dbl_dblP > dblLowerClassBoundary - ((dblLowerClassBoundary - dblUpperClassBoundary) * 0.2) And dbl_dblP < dblLowerClassBoundary - ((dblLowerClassBoundary - dblUpperClassBoundary) * 0.03) Then
    strRAG = "33023" '"&H000080FF&"
Else
    strRAG = "65280" '"&H0000FF00&"
End If

ReturnRAG_Colour = strRAG
End Function
Function ReturnHighFor1(ByVal intClass_ID As Integer) As String
Select Case intClass_ID
    Case 1
        ReturnHighFor1 = "High"
    Case 2
        ReturnHighFor1 = "Good"
    Case 3
        ReturnHighFor1 = "Moderate"
    Case 4
        ReturnHighFor1 = "Poor"
    Case 5
        ReturnHighFor1 = "Bad"
        
End Select
End Function
