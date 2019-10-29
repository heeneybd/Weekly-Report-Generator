# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import pandas as pd
import random

from pandas import ExcelWriter
from pandas import ExcelFile

path = r"C:\Users\braden.h\Desktop\rabbit_class.xlsx"



df = pd.read_excel(path, 'week4')
df1 = pd.read_excel(path, 'student_names')

behavior1 = [" was very well behaved this week.", " has had good behavior this week.", " was a perfect little angel this week."]
behavior2 = [" was fairly well behaved this week.", " did well in class this week.", " was just perfectly average this week."]
behavior3 = [" was not very well behaved this week.", " showed poor behavior this week.", " was a little shit this week."]

speaking1 = [" is able to speak quite well.", " has very good speaking skills and is able to make the sounds well.", " can speak well for this level."]
speaking2 = [" is able to speak fairly well.", "'s speaking continues to show steady improvement.", " is able to say words and speak fairly well."]
speaking3 = ["'s speaking needs improvement.", " still needs more practice speaking.", " needs more practice speaking loudly and clearly."]

listening1 = [" has excellent listening skills and is able to clearly understand spoken English.", "'s listening skill is quite good.", " has excellent listening comprehension."]
listening2 = ["'s listening skills continue to show improvement.", "'s listening skills are at the appropriate level.", " has no problem understanding spoken English."]
listening3 = ["'s listening skills need improvement.", " needs more practice with listening.", " has difficulty understand spoken English but we are still working on improving this skill."]

print("Rabbit Class:")
print('')

student_list = df1['Student Name'].tolist()
student_pronounlist = df1['Pronoun'].tolist()

i = 0
for i in df.index:
    student_number = df["Student Number"][i]

    student_name = student_list[int(student_number) - 1]
    
    student_pronoun = student_pronounlist[int(student_number) - 1]
    
    
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
    
    paragraph = [listeningstring, speakingstring, behaviorstring]

    random.shuffle(paragraph)
    
    paragraph[0] = student_name + paragraph[0]
    paragraph[1] = student_pronoun + paragraph[1]
    paragraph[2] = student_pronoun + paragraph[2]
    
    paragraph = ' '.join(paragraph) 
    
    paragraph = paragraph.replace("She's", "Her")
    paragraph = paragraph.replace("He's", "His")
    
    print(paragraph)
    print('')