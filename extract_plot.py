import Analyze
import System, System.IO

################### COPYRIGHT FIKRIMF & GEMA @IMMOBI TECHNOLOGY SDN BHD###############

#exportfolder = "E:\\Export Plot DT\\UMOBILE - NOKIA\\KUAT\\"        #TARGET FOLDER .JPG/.PNG
#workbookname = 'PAC'

#ExportTargetPath = "E:\\Export Plot DT\\Q2 ALL HIGHWAY\\GUTHRIE\\TAB\\"               #TARGET FOLDER .TAB
#ExportTargetName = 'Q2 ALL HIGHWAY GUTHRIE'

warningtext = 'Please wait until the workbook finisihed loading' 
Windows.MessageBox(warningtext, 'Loading the logfiles...', 'yes')

#if not System.IO.Directory.Exists( exportfolder ):
#	System.IO.Directory.CreateDirectory( exportfolder )

Width = 260 #280

# Set workspace's DataViewAdded-event handler.
#Analyze.Workspace.DataViewAdded += lambda sender, args: dataviewadded(args.dataview)
Longitude = float(103.0851)
Latitude = float(5.3364) 
# get all workbooks check type and set width
index=0
workbooks = Analyze.Workspace.Workbooks
for workbook in workbooks:
	#filename = '{0} {1}'.format(workbookname,index)
	pages = workbook.GetAllPages()
	index2=0
	for page in pages:	
		dataviews = page.Views
		namaPage = page.Title 
		for dataview in dataviews:
			if isinstance(dataview, MapDataView):
				dataview.ResizeSidepanel(Width)
				dataview.BestFitZoom()
				#dataview.ZoomTo(abcde)
				dataview.SetDrawingMode("line")
				dataview.SetShowScaleBar(False)
				dataview.ResizeSidepanel(Width)
				frames = dataview.Frames
				for frame in frames:
					frame.ResizeFrame(200,0,270,150)
				#workbook.ExportPage(index2,exportfolder+str(index2)+str(workbook)+"_"+workbookname+"_"+namaPage+".jpg")
				#workbook.ExportToWord(exportfolder, "A.docx",'','')
			    #Log.Write(namaPage+" "+"successfully exported !!!")			    
		index2=index2+1
	index=index+1
	#workbook.ExportToWord(exportfolder, "A.docx",'','')
#zzzz = dataview.GetCenter()
#bv = str(zzzz)
#Windows.MessageBox(bv, 'campret', 'as')
#################################################### EXPORT MAP TO TAB #######################################################
# Creates MapInfo (.tab) files from all layers in all map views and export files to given folder. 
#def exportMapsFromWorkboobPage(page, index):
#	dataviews = page.Views
#	datahalaman = page.Title
#	for dataview in dataviews:
#		# check that dataview is MapDataView type
#		if isinstance(dataview, MapDataView):
#			filename = '{1} {0}'.format(index,ExportTargetName+" "+datahalaman)
#			pathAndFileName = '{0} {1}'.format(ExportTargetPath,filename)
#			dataview.ExportToMapInfo(pathAndFileName,'')
#			index = index +1
#	return index

## Create the folder if it does not exist.
#if not System.IO.Directory.Exists( ExportTargetPath ):
#	System.IO.Directory.CreateDirectory( ExportTargetPath )

## Get all workbooks and go through all pages.
#workbooks = Analyze.Workspace.Workbooks
#index = 1
#for workbook in workbooks:
#	workbookpageindex = 0
#	workbookpages = workbook.GetAllPages()
#	for workbookpage in workbookpages:
#		index = exportMapsFromWorkboobPage(workbookpage, index)