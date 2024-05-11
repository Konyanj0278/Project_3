import pandas as pd
import argparse
import os
import pymongo
import ffmpeg
import sys
import xlsxwriter
#Argparse setup
#fio-u-pRmOl_Jz1K25ztWSDW0HBmZlSoBpdKVQgffEuL6Uf1DDOEUC-vAIMIRa4x4exV1r
argparser = argparse.ArgumentParser(description=
'''This script will take a Baselight export file and a Xytech file 
and uploaded the results to a DB. The script will also process video files before exporting
them to a XML file.''')
argparser.add_argument('--baselight', help='The Baselight export file')
argparser.add_argument('--xytech', help='The Xytech file')
argparser.add_argument('--process', help='Process video files')
argparser.add_argument('--output', help='Export to XML file')
argparser.add_argument('--thumbnail', help='Generate thumbnail')

#MongoDB setup
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["File_Exports"]
Baselight_col = mydb["baselight"]
Xytech_col = mydb["xytech"]

# Global Variables
Video_Max_Frames = 0
video_stream = None

#Populate DB with Baselight data

def populate_db_baselight(baselight_file):
    count = 0
    file=open(baselight_file, "r")
    
    BL_File = file.read().splitlines()
    for techfile in BL_File:
        count += 1
        file_dict = {}
        max_frames = 0
        try:
            file_dict['Folder'] = techfile.split(' ',1)[0]
            frames = techfile.split(' ',1)[1]
            file_dict['Frames'] = frames
        except:
            return
        
        for frame in frames.split(' '):
            try:
                max_frames = max(max_frames, int(frame))
            except:
                continue
   
        file_dict['Max Frames'] = max_frames
        Baselight_col.insert_one(file_dict)
    
#Populate DB with Xytech data

def populate_db_xytech(xytech_file):
    count = 0
    file=open(xytech_file, "r")
    workorder= {}
    XY_File = file.read().splitlines()
    read_notes = False
    
    for techfile in XY_File:
        count += 1

        location= {}
        if 'Producer' in techfile:
            workorder['Producer'] = techfile.split(':')[1]
        if 'Operator' in techfile:
            workorder['Operator'] = techfile.split(':')[1]
        if 'Job' in techfile:
            workorder['Job'] = techfile.split(':')[1]
       
        if read_notes:
            print(techfile)
            workorder['Notes'] = techfile
            read_notes = False
        if 'Notes' in techfile: 
            read_notes = True 

        if '/' in techfile:
            location['Location'] = techfile
            Xytech_col.insert_one(location)

        print(workorder)
    Xytech_col.insert_one(workorder)
    



def get_video_info(in_filename):
    global Video_Max_Frames
    try:
        probe = ffmpeg.probe(in_filename)
  
        
    except ffmpeg.Error as e:
        print('Error occurred', file=sys.stderr)
    
    video_stream = next((stream for stream in probe['streams'] if stream['codec_type'] == 'video'), None)
    if video_stream is None:
        print('No video stream found', file=sys.stderr)
        sys.exit(1)
    Video_Max_Frames = video_stream['nb_frames']
    return video_stream

def generate_thumbnail(filename,in_filename, frame_time):
 
    out_filename = filename + ".png"
    
    try:
        
        ffmpeg.input(in_filename, ss=frame_time).output(out_filename,loglevel="quiet", vframes=1,s="96x74").overwrite_output().run()
    except ffmpeg.Error as e:
       
        print("Error")
        sys.exit(1)


    return out_filename

def convert_frame_to_time(frame,full_timecode,export):
        fps = video_stream['r_frame_rate'].split('/')[0]
        fps = int(fps)
       
        minutes = 0
        hours = 0
        seconds = int(frame) / fps
        frames = int(frame) % fps
        if seconds > 60:
            minutes = seconds / 60
            seconds = seconds % 60
        if minutes > 60:
            hours = minutes / 60
            minutes = minutes % 60
        result = "{0:02d}:{1:02d}:{2:02d}".format(int(hours), int(minutes), int(seconds))
        if (full_timecode == True and export == False):

            timecode = result + "." + str(frames)
            return timecode
        if (full_timecode == True and export == True):
            timecode = result + ":" + str(frames)
            return timecode
        else:
            return result
def generate_video_clip(in_filename, out_filename, start_time, end_time):
    global video_stream
    fps = video_stream['r_frame_rate'].split('/')[0]
    
    estart_time = int(start_time) / int(fps)
    eend_time = int(end_time) / int(fps) - estart_time
    print("start time: ", estart_time)
    print("end time: ", eend_time)

    try:
        input_file = ffmpeg.input(in_filename,ss=estart_time,t=eend_time).output((out_filename + ".mp4"),loglevel='quiet').overwrite_output().run()
    except ffmpeg.Error as e:
        print("Error")
        print(e.stderr)
        sys.exit(1)
    return
def upload_to_frameio(filename):
    from frameioclient import FrameioClient

    client = FrameioClient("fio-u-pRmOl_Jz1K25ztWSDW0HBmZlSoBpdKVQgffEuL6Uf1DDOEUC-vAIMIRa4x4exV1r")

    asset = client.assets.upload(
        destination_id="cbeec5e6-d995-4a77-a043-bcdbeb0ecd28",
        filepath=filename
        )

def process_frames(video,BL_File, XY_File, in_filename):
    columns = ['Producer','Operator','Job','Notes']
    data =[]
    producer='Producer'
    operator='Operator'
    Job='Job'
    Notes='Notes'
    with open("./Xytech.txt") as f: 
        XY_File = f.read().splitlines()

    
    for techfile in XY_File:

        if(producer in techfile):
            producer=techfile.split(':')[1]
        if(operator in techfile):
            operator=techfile.split(':')[1]
        if(Job in techfile):
            Job= techfile.split(':')[1]
        else:
            Notes = techfile

    data.append((producer,operator,Job,Notes))
    data.append(('','','',''))
    data.append(('Location','Frames to Fix', '',''))
    count = 0
    for currentReadLine in BL_File:

        parseLine = currentReadLine.split()

        currentFolder = parseLine.pop(0)
        parseFolder = currentFolder.split("/")
        parseFolder.pop(1)
        newFolder = "/".join(parseFolder) 

        for techfile in XY_File:
            if newFolder in techfile:
                currentFolder = techfile.strip()

        tempStart=0
        tempLast = 0


        for number in parseLine:

            if not number.isnumeric():
                continue
            if tempStart == 0:
                tempStart = number
                continue

            if number == str(int(tempStart)+1):
                tempLast=number
                continue

            elif number == str(int(tempLast)+1):
                tempLast = number
                continue

            else:
                if int(tempLast) > 0:
                    frame_start=convert_frame_to_time(tempStart, True, True)
                    frame_end=convert_frame_to_time(tempLast, True, True)
                    middle_frame = (int(tempStart) + int(tempLast)) / 2
                    middle_frame = convert_frame_to_time(middle_frame, True, False)
                    thumbnail=generate_thumbnail(filename="Row "+str(count),in_filename=video, frame_time=middle_frame)
                    generate_video_clip(in_filename, "Clip "+str(count), tempStart, tempLast)
                    data.append((currentFolder,str(tempStart) + "-" + str(tempLast),str(frame_start)+ '-'+frame_end,))
                    count += 1
            
                
                tempStart = number
                tempLast = 0
                
        if int(tempLast) > 0:
            frame_start=convert_frame_to_time(tempStart, True, True)
            frame_end=convert_frame_to_time(tempLast, True, True)
            middle_frame = (int(tempStart) + int(tempLast)) / 2
            middle_frame = convert_frame_to_time(middle_frame, True, False)
            thumbnail=generate_thumbnail(filename="Row "+str(count),in_filename=video, frame_time=middle_frame)
            generate_video_clip(in_filename, "Clip "+str(count),tempStart, tempLast)
            data.append((currentFolder,str(tempStart) + "-" + str(tempLast),str(frame_start)+ '-'+frame_end,''))
            count += 1
        
    df = pd.DataFrame(data, columns=columns)

    workbook =pd.ExcelWriter('Project.xlsx', engine='xlsxwriter')
    df.to_excel(workbook, sheet_name='Sheet1')
    #make loop for thumbnails from g3 onwards
    worksheet = workbook.sheets['Sheet1']
    
    for i in range(len(data)):
        thumbnail = 'Row '+str(i)+'.png'
        index = 'G'+str(i+5)
        if (i+2) >(len(data)-2):
            break
        worksheet.insert_image(index, thumbnail, {'x_scale': 0.5, 'y_scale': 0.25})
    workbook.close()
    for i in range(len(data)):
        if (i+2) >(len(data)-2):
            break
        vid = 'Clip '+str(i)+'.mp4'
        upload_to_frameio(vid)



def process_video_files(in_filename, time):
    global video_stream
    BL_file =[]
    video_stream = get_video_info(in_filename)
    frames_time=convert_frame_to_time(time, False, False)
    print("Frames Time: ", frames_time)
  
    print("Max Frames: ", Video_Max_Frames)
    db_query = {"Max Frames": {"$lt": int(Video_Max_Frames)}}
    BL_cursor = Baselight_col.find(db_query)
    for document in BL_cursor:
        BL_file.append(document['Folder'] + " " + document['Frames'])
    XY_cursor = Xytech_col.find({})
    XY_file = []
    for document in XY_cursor:
        XY_file.append(document)

    process_frames(in_filename,BL_file, XY_file, in_filename)
    return

if argparser.parse_args().baselight:
    populate_db_baselight(argparser.parse_args().baselight)

if argparser.parse_args().xytech:
    populate_db_xytech(argparser.parse_args().xytech)


if argparser.parse_args().process:
    process_video_files(argparser.parse_args().process, 10)

if argparser.parse_args().thumbnail:
    image=generate_thumbnail("test",argparser.parse_args().thumbnail, 5166)
    

'''





columns = ['Producer','Operator','Job','Notes']
data =[]
producer='Producer'
operator='Operator'
Job='Job'
Notes='Notes'
with open("./Xytech.txt") as f: 
	XY_File = f.read().splitlines()

BL_File = open("./Baselight_export.txt", "r") 
for techfile in XY_File:
		
		if(producer in techfile):
			producer=techfile.split(':')[1]
		if(operator in techfile):
			operator=techfile.split(':')[1]
		if(Job in techfile):
			Job= techfile.split(':')[1]
		else:
			Notes = techfile

data.append((producer,operator,Job,Notes))
data.append(('','','',''))
data.append(('Location','Frames to Fix', '',''))

for currentReadLine in BL_File:

	parseLine = currentReadLine.split()

	currentFolder = parseLine.pop(0)
	parseFolder = currentFolder.split("/")
	parseFolder.pop(1)
	newFolder = "/".join(parseFolder) 
	
	for techfile in XY_File:
		if newFolder in techfile:
			currentFolder = techfile.strip()
		
	tempStart=0
	tempLast = 0
	
	
	for number in parseLine:
		
		if not number.isnumeric():
			continue
		if tempStart == 0:
			tempStart = number
			continue

		if number == str(int(tempStart)+1):
			tempLast=number
			continue

		elif number == str(int(tempLast)+1):
			tempLast = number
			continue

		else:
			if int(tempLast) > 0:
				data.append((currentFolder,str(tempStart) + "-" + str(tempLast),'',''))
			else:
				data.append((currentFolder,str(tempStart) ,'',''))
			tempStart = number
			tempLast = 0
	if int(tempLast) > 0:

		data.append((currentFolder,str(tempStart) + "-" + str(tempLast),'',''))
	else:
		data.append((currentFolder,str(tempStart) ,'',''))
		tempStart = number
		tempLast=0
df = pd.DataFrame(data, columns=columns)

df.to_csv('Project1.csv', encoding='utf-8', index=False)


'''