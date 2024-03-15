# using flask_restful 
from flask import Flask, jsonify, request 
from flask_restful import Resource, Api 
import joblib
from sklearn.feature_extraction.text import TfidfVectorizer
import pandas as pd
import pyautogui
import time
import os
from bs4 import BeautifulSoup
from html2image import Html2Image


app = Flask(__name__) 

app = Api(app) 
class Model(Resource): 

	def get(self): 
		 
            
         data=pd.read_excel('FIATMARK.xlsx')
         columns=data.columns
         filenames=[]
         contact=[]
         cwd=os.getcwd()
         table_columns=columns[2:]
         hti = Html2Image(size=(800, 1000))
         for i in range(5):
                
            with open("template.html") as file:  
                template = file.read() 
                        
                soup = BeautifulSoup(template, 'html.parser')
                name_tag=soup.find_all(id="name")[0]
                roll_no_tag=soup.find_all(id="roll_number")[0]
                name_tag.string="Name : "+data.loc[i,'Name']
                roll_no_tag.string="Register Number : "+str(data.loc[i,'Register Number'] )
                tag = soup.find_all(class_='grade-report')[0]
            
                for col in table_columns:
                    new_tag1=soup.new_tag('td',id="",class_="")
                    new_tag2=soup.new_tag('td',id="",class_="")
                    par=soup.new_tag('tr',class_="append")
                    new_tag1.string=col
                    new_tag2.string= str(data.loc[i,col])
                    par.append(new_tag1)
                    par.append(new_tag2)
                    tag.append(par)
                    
                modified_html=soup.prettify()
                roll_no=data.loc[i,'Register Number']
                hti.screenshot(
                    html_str=modified_html, 
                    css_str='',
                    save_as=f'{roll_no}.png'
                )
                filenames.append(f'{roll_no}.png')
                #contact.append(data.loc[i,'Contact'])
                contact.append("9360670658")
                
         time.sleep(1)
         pyautogui.press('win') 
         time.sleep(1)
         pyautogui.write('whatsapp') 
         time.sleep(1)
         pyautogui.press('enter')  
         time.sleep(1.5)
         pyautogui.press('search')

         for number,file in zip(contact,filenames):
            time.sleep(0.5)
            pyautogui.write(number) 
            time.sleep(0.5) 
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)  
            image_path = os.path.join(cwd,file)
            pyautogui.typewrite(' ')
            time.sleep(0.5)
            pyautogui.hotkey('shift','tab')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.write(image_path)
            time.sleep(0.5)
            pyautogui.press('enter')    
            time.sleep(0.5) 
            pyautogui.press('enter') 
            time.sleep(0.5)

            for i in range(5):
                pyautogui.press('tab')
                time.sleep(0.5)
         return " hello ";



api.add_resource(Model, '/model') 


if __name__ == '__main__': 

	app.run(debug = True) 
