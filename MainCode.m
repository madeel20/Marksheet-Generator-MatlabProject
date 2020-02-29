
clc % Clear command window.
clear all %Clear variables and functions from memory
close all % closes all the open figure windows
home % Send the cursor home

filename = 'Registration_Details.xls';
[num,txt] = xlsread(filename)
% Read Excel sheet containing details. Text is read from the file
% seperately from numbers


len=length(txt)
len = 3
% obtain length of text in excel or number of certificates to be generated
% This code provides scalability

blankimage = imread('Certificate_Blank.tif');
% Read blank certificate image

for i=1:len
    for j= 2:2
        text_names(i,j)=txt(i,j)
    end
end
% Obtain names from the txt variable which are in 2nd column

for i=1:len
    for j= 3:3
      text_date(i,j)=txt(i,j)
    end
end
% Obtain dates from the txt variable which are in 3rd column
for i=1:len
    for j= 4:4
      text_rollno(i,j)=txt(i,j)
    end
end
% Obtain roll no from the txt variable which are in 3rd column
for i=1:len
    for j= 5:5
      text_enroll(i,j)=txt(i,j)
    end
end
% Obtain enrollment from the txt variable which are in 3rd column
for i=1:len
    for j= 6:6
      text_batch(i,j)=txt(i,j)
    end
end
% Obtain Batch from the txt variable which are in 3rd column
for i=1:len
    for j= 7:7
      text_fname(i,j)=txt(i,j)
    end
end
% Obtain Batch from the txt variable which are in 3rd column

%Ignore first row which is heading
for i=2:len
        text_all=[text_names(i,2) text_date(i,3) text_rollno(i,4) text_enroll(i,5) text_batch(i,6) text_fname(i,7)]
        % combine names and topics
        
        position = [900 455;1760 400;940 555;1400 555;1800 555;500 505];         
        % obtain positions to insert on image, MSPaint or any image editor
        
        RGB = insertText(blankimage,position,text_all,'FontSize',22,'BoxOpacity',0);
        %Provide parameters for function insertText
        %Font size is 22 and opacity of box is 0 means 100% transparent
        
        figure
        imshow(RGB)        
        %Obtain and display figure in color
        
        y=i-1
        filename=['Certificate_Topic_' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
        % generate and save files with .tif extension
end    

% Code works for any length of data ROWWISE AND COLUMNWISE
% Last for loop with minor changes facilitate new project implemntations
