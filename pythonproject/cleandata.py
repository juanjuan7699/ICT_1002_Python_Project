import iplot as iplot
import pandas as pd
import DateTime

#get data from github
url = 'https://github.com/owid/covid-19-data/raw/master/public/data/owid-covid-data.xlsx'
df = pd.read_excel(url, header=0)


# onl include continent from Asia to do analysis in ASEAN countries.
asia=df[df['continent']=='Asia']
asia = asia[(asia.location.isin(["Brunei","Indonesia","Cambodia","Laos","Malaysia","Myanmar","Philippines","Singapore","Thailand","Timor","Vietnam"]))]
#removing unnessary columns
asia.drop(asia.columns[[6, 9, 10,11,12,13,14,15,18,19,20,21,22,24]], axis = 1, inplace = True)
#replacing NaN with 0
asia.fillna(0, inplace=True)

#group dataset into weekly basis
asia['date'] = pd.to_datetime(asia['date'])
asia['Week_Number'] = asia['date'].dt.week
asean = asia.groupby(['iso_code','location','Week_Number'],as_index=False).agg({'total_cases':'max','new_cases':"sum",'total_deaths':'max','new_deaths':'sum',
                                                    'new_tests':'sum','total_tests':'max','positive_rate':'mean','stringency_index':'mean',
                                                    'population':'last'	,'population_density':'mean','median_age':'mean','aged_65_older':'mean',
                                                    'aged_70_older':'mean','gdp_per_capita':'mean','extreme_poverty':'mean','cardiovasc_death_rate':'mean',
                                                    'diabetes_prevalence':'mean','female_smokers':'mean','male_smokers':'mean','handwashing_facilities':'mean',
                                                    'hospital_beds_per_thousand':'mean','life_expectancy':'mean','human_development_index':'mean'})
asean=asean.sort_values(by=['Week_Number'])

#export to excel
asean.to_excel("covidinasean.xlsx",index=False)


