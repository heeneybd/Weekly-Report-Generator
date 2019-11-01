# -*- coding: utf-8 -*-
"""
Spyder Editor
This is a temporary script file.
"""

import os
import pandas as pd
import random
import xlsxwriter
from datetime import date
import calendar 

##This is the path to the excel file with raw data
path = r"C:\Users\braden.h\Desktop\rabbit_class.xlsx"

##These are the worksheet names in the excel fle
df = pd.read_excel(path, 'week4')
df1 = pd.read_excel(path, 'student_names')

##These are the possible phrases that could be output (without names or pronouns)
behavior1 = [" was very well behaved this week.", " has had good behavior this week.", " was a perfect little angel this week.", " is very well-behaved and sets a positive example for other students in the class.", " had excellent behavior this week and followed directions well."]
behavior2 = [" was fairly well behaved this week.", " did well in class this week.", " was pretty well-behaved this week.", " followed directions well this week and was fairly well-behaved.", " did well in class this week."]
behavior3 = [" was not very well behaved this week.", " showed poor behavior this week.", " did not follow directions well this week.", " needs to work harder to follow the class rules.", " doesn't always listen to the teacher and follow the rules."]
speaking1 = [" is able to speak quite well.", " has very good speaking skills and is able to make the sounds well.", " can speak well for this level.", " is able to speak English loudly and clearly.", " has excellent speaking skills."]
speaking2 = [" is able to speak fairly well.", "'s speaking continues to show steady improvement.", " is able to say words and speak fairly well.", "'s speaking ability continues to show progress and improvement.", " is getting more comfortable using everyday English."]
speaking3 = ["'s speaking could use practice.", " still needs more practice speaking.", " needs more practice speaking loudly and clearly.", "'s speaking still needs more practice.", " lacks confidence when speaking English.", "'s speaking skill needs extra practice."]
listening1 = [" has excellent listening skills and is able to clearly understand spoken English.", "'s listening skill is quite good.", " has excellent listening comprehension.", " has good listening skills and can understand English well for this level.", "'s listening ability is quite good."]
listening2 = ["'s listening skills continue to show improvement.", "'s listening skills are at the appropriate level.", " has no problem understanding spoken English.", " has no difficulty understanding basic English.", "'s listening ability continues to show improvement."]
listening3 = ["'s listening skills need improvement.", " needs more practice with listening.", " has difficulty understand spoken English.", "'s listening skills still need improvement.", " needs more practice listening for English comprehension."]

##These lists qualifies the different strings as having positive or negative value 
##so the program knows to combine using "but" or "and" when constructing the comments.
positive = behavior1 + behavior2 + speaking1 + speaking2 + listening1 + listening2
negative = behavior3 + speaking3 + listening3

print("Rabbit Class:")
print('')

student_list = df1['Student Name'].tolist()
student_pronounlist = df1['Pronoun'].tolist()

list = []
i = 0
for i in df.index:
    student_number = df["Student Number"][i]
    student_name = student_list[int(student_number) - 1]
    student_pronoun = student_pronounlist[int(student_number) - 1]
    
##This code generates random sentences correlating to their scores in each category.
    behavior = df["Behavior"][i]
    if behavior == 1:
        behaviorstring = random.choice(behavior1)
    elif behavior == 2:
        behaviorstring = random.choice(behavior2)
    else:
        behaviorstring = random.choice(behavior3)
    
    speaking = df["Speaking"][i]
    if speaking == 1:
        speakingstring = random.choice(speaking1)
    elif speaking == 2:
        speakingstring = random.choice(speaking2)
    else:
        speakingstring = random.choice(speaking3)
    
    listening = df["Listening"][i]
    if listening == 1:
        listeningstring = random.choice(listening1)
    elif listening == 2:
        listeningstring = random.choice(listening2)
    else:
        listeningstring = random.choice(listening3)

##this changes the order of the sentences in the listening_speaking 
##and combines sentences 2 and 3 with "and" or "but" depending on whether
##they qualify as positive or negative        
    listening_speaking = [listeningstring, speakingstring]
    random.shuffle(listening_speaking)
    
    if listening_speaking[0] in positive:
        if listening_speaking[1] in positive:
            listening_speaking[0] = listening_speaking[0][:-1] + " and"
            
        elif listening_speaking[1] in negative:
            listening_speaking[0] = listening_speaking[0][:-1] + " but"
    elif listening_speaking[0] in negative:
        if listening_speaking[1] in positive:
            listening_speaking[0] = listening_speaking[0][:-1] + " but"
        elif listening_speaking[1] in negative:
            listening_speaking[0] = listening_speaking[0][:-1] + " and"
    behaviorstring = student_name + behaviorstring + "  "
    listening_speaking[0] = student_pronoun + listening_speaking[0]

    listening_speaking[1] = student_pronoun + listening_speaking[1]
    listening_speaking[1] = listening_speaking[1][0].lower() + listening_speaking[1][1:]
    

   
    listening_speaking = ' '.join(listening_speaking)   
    listening_speaking = listening_speaking.replace("She's", "Her")
    listening_speaking = listening_speaking.replace("He's", "His")
    listening_speaking = listening_speaking.replace("she's", "her")
    listening_speaking = listening_speaking.replace("he's", "his")
    
    print(behaviorstring + listening_speaking)
    print('')
    list.append(behaviorstring + listening_speaking)
    
df = pd.DataFrame({'Rabbit Class Comments': list})
writer = pd.ExcelWriter(r'C:\Users\braden.h\Desktop\ %s comments.xlsx' % date.today(), engine='xlsxwriter')
df.to_excel(writer, sheet_name='Weekly Comments')
writer.save()