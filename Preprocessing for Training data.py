file_name="World War II.docx"

import pandas as pd
import textract

text=textract.process(file_name)
text=text.decode("utf-8")

lines=text.splitlines()

df=pd.DataFrame()

df['lines']=lines
df['char_length']=df['lines'].apply(lambda x : len(x))
df['no_of_words']=df['lines'].apply(lambda x: len(x.split()))

df['index_']=df.index

#" Allots Values for Space"

def emptylines(df):
    for i in range(len(df)):
        if(i==0):
            if(df.loc[i+1]["lines"]==""):
                df.at[i,"space"]=1
            else:
                df.at[i,'space']=0
            i+=1
            continue

        if(i==len(df)-1):
            if(df.loc[i-1]["lines"]==""):
                df.at[i,"space"]=1
            else:
                df.at[i,'space']=0
            i+=1
            continue

        if((df.loc[i-1]['lines']=="") ^(df.loc[i+1]['lines']=="")):
           df.at[i,'space']=1
        elif((df.loc[i-1]['lines']=="") and (df.loc[i+1]['lines']=="")):
           df.at[i,'space']=2
        else:
           df.at[i,'space']=0
        i+=1

def caps(text):
    caps_count=0
    for i in text:
        if(i.isupper()):
            caps_count+=1
    if caps_count==0 or len(text)==0:
        return 0
    else:
        return  caps_count/len(text)
    
def feature_generator(df):
    len_df=len(df)
    for i in range(len(df)):
        if(i==0):
            df.at[i,"next_char"]=df.loc[i+1]["char_length"]
            df.at[i,"next_words"]=df.loc[i+1]["no_of_words"]
            df.at[i,"next_next_char"]=df.loc[i+2]["char_length"]
            df.at[i,"next_next_words"]=df.loc[i+2]["no_of_words"]
            df.at[i,"next_caps"]=caps(df.at[i+1,"lines"])
            df.at[i,"prev_char"]=0
            df.at[i,"prev_words"]=0
            df.at[i,"prev_prev_char"]=0
            df.at[i,"prev_prev_words"]=0
            i+=1
            continue
            
        if(i==1):
            df.at[i,"next_char"]=df.loc[i+1]["char_length"]
            df.at[i,"next_words"]=df.loc[i+1]["no_of_words"]
            df.at[i,"next_next_char"]=df.loc[i+2]["char_length"]
            df.at[i,"next_next_words"]=df.loc[i+2]["no_of_words"]
            df.at[i,"prev_char"]=df.loc[i-1]["char_length"]
            df.at[i,"prev_words"]=df.loc[i-1]["no_of_words"]
            df.at[i,"prev_prev_char"]=0
            df.at[i,"prev_prev_words"]=0      
            
            df.at[i,"next_caps"]=caps(df.at[i+1,"lines"])
            i+=1
            continue
            
        if(i==len_df-2):
            df.at[i,"next_char"]=df.loc[i+1]["char_length"]
            df.at[i,"next_words"]=df.loc[i+1]["no_of_words"]
            df.at[i,"next_next_char"]=0
            df.at[i,"next_next_words"]=0
            df.at[i,"prev_char"]=df.loc[i-1]["char_length"]
            df.at[i,"prev_words"]=df.loc[i-1]["no_of_words"]
            df.at[i,"prev_prev_char"]=df.loc[i-2]["char_length"]
            df.at[i,"prev_prev_words"]=df.loc[i-2]["no_of_words"]
            
            df.at[i,"next_caps"]=caps(df.at[i+1,"lines"])
            i+=1
            continue
            
            
        if(i==len_df-1):
            df.at[i,"next_char"]=0
            df.at[i,"next_words"]=0
            df.at[i,"next_next_char"]=0
            df.at[i,"next_next_words"]=0
            df.at[i,"prev_char"]=df.loc[i-1]["char_length"]
            df.at[i,"prev_words"]=df.loc[i-1]["no_of_words"]
            df.at[i,"prev_prev_char"]=df.loc[i-2]["char_length"]
            df.at[i,"prev_prev_words"]=df.loc[i-2]["no_of_words"]
            
            df.at[i,"next_caps"]=0
            i+=1
            continue
        
        df.at[i,"next_char"]=df.loc[i+1]["char_length"]
        df.at[i,"next_words"]=df.loc[i+1]["no_of_words"]
        df.at[i,"next_next_char"]=df.loc[i+2]["char_length"]
        df.at[i,"next_next_words"]=df.loc[i+2]["no_of_words"]
        df.at[i,"prev_char"]=df.loc[i-1]["char_length"]
        df.at[i,"prev_words"]=df.loc[i-1]["no_of_words"]
        df.at[i,"prev_prev_char"]=df.loc[i-2]["char_length"]
        df.at[i,"prev_prev_words"]=df.loc[i-2]["no_of_words"]
        
        df.at[i,"next_caps"]=caps(df.at[i+1,"lines"])
        i+=1
            

df['caps']=df['lines'].apply(caps)
emptylines(df)
feature_generator(df)
df["label"]=0
df['label2']=0
#Section - 0  Heading -1

df.to_excel(file_name.split(".")[0]+"_data.xlsx")

