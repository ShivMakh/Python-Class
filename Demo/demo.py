# -*- coding: utf-8 -*-
"""
@author: ShivMakh

How to read this script?

like any other doc

start at the top and read down line by line and let the info build
on the previous discussions



Comments - with triple quotes or with a pund/hastag(#)
These explain what is happening and why

most of the comments i have are business level, but plenty of technical
aspects that glossed over


"""


"""

importing some packages

packages allow us to advantage of different features and tools with python
without having to create it from scratch


we should normally import packages at the start of the script so that your
toolbox is full

"""

# packages work with numbers and create tables
# these are standard imports, and effectively recreate excel features

# visualizations packages

# createa and send emails on outlook


# we create the main namespace
# this is a best practice thing,
# you can do without it, however allows for more consistent runs across machines


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import win32com.client as win32
if __name__ == "__main__":

    '''
    we are almost ready to start coding

    general notes to keep in mind

        1. this is meant to get you started, so we are going to do the manual
        process step by step

        2. most of the time translating is pretty straightforward, once you
        get used to it. its mostly similar, but there is just some minor
        differences and the comments will walk you through some of it

        just know when something that isnt directly or intuitvely similar to
        excel functionality,

        the first few results on search engines usually have the soln you need


        goal is getting you comfortable using online resources, not to becoming
        a pro programmer

        that being said - dont copy my code into internet and code from the
        internet into your code. this is doubly true with AI generated code

        AI generated code is often much much more complicated and harder to
        understand and decode when things go wrong/things are not working as
        you would want.

        search things like
            'python pivot table?', 'read execl into python',
            'how to make pie chart in python'


        3. code in my script is not the best way to do things, but it
        is the code that i understand and that is more important than optimized
        soln

    '''

    # first set some default paths so that it can be dynamic
    # python needs to know where to look for input data
    # update with your own file path

    data_folder = r'C:\Users\i5 PC\Documents\GitHub\Python-Class\Demo'

    # not demo specific, however very useful python features
    # setting the date as a variable so if you need to use it

    today = pd.to_datetime('today').date()
    specific_date = pd.to_datetime('12/5/2025')

    '''Lets start recreating the excel steps that we have in our
    Super Store data sample'''

    # ---- Step 1
    # directions say to make sure that the tables have the correct regional
    # managers
    # a table in python is called a dataframe

    regional_managers_data = {'West': 'Sadie Pawthorne',
                            'East': 'Chuck Magee',
                            'Central': 'Rox Rodriguez',
                            'South': 'Fred Suzuki'
                            }

    regional_managers = pd.DataFrame(regional_managers_data.items(),
                                    columns=['Region', 'Managers'])

    # ---- Step 2

    # now copy and paste the data set into the template
    # that is to get all the data into the same location to generate the report 
    # more easily
    # another way of thinking about this is read all the excel data into the
    #same location

    # i use the data_folder that we created before. by doing that first i only
    # need to change that line of code as file
    # path changes, generally you dont want to consider things to be static,
    # so creating dynamic variables makes
    # some updates easier, as you go through the process you will learn how to
    #ID these things and where this set up makes sense
    orders_dataset_df = pd.read_excel(fr'{data_folder}\Superstore Dataset.xlsx')
    #df is a common acroynm, and its a the same as a excel table
    return_status_df = pd.read_excel(fr'{data_folder}\Return Status.xlsx')
    
    #directions warns about duplicated rows, from returns
    return_status_df = return_status_df.drop_duplicates()
    
    #---- Step 3
    
    #third step is to make sure that the df is the excel formulas work properly
    
    #in python you can ort columes and the process si similar, however it is 
    # not needed for combining the dataset like we do in excel
    
    #will provide the code here as its helpful
    
    # i comment out the line here so it does not actually overwrite anything,
    # however it allows me to show you the syntax 
    # commenting code like this is nice as you develop code and are trying to do
    # stuff with and without certain steps
    
    # return_status_df = return_status_df.sort_values(by='Order ID'. ascending=True)
    

    #---- Step 4
    
    # the next step is to update the formulas, in this case we cannot update 
    # the formulas
    # we need to actually program the forumlas
    
    # the first formula is getting the correspidning regional manager that is
    # responsible with the order inthe same table
    # we want to "merge" two different tables based on some common characterisitc
    # in excel this is often achieved by vlookup,hlookup,xlookup
    
    orders_dataset_managers = pd.merge(orders_dataset_df, regional_managers, on='Region', how='left')
    
    #now we want to "merge" this data with the return status
    
    orders_dataset= pd.merge(orders_dataset_managers, return_status_df, on=['Order ID'], how='left')
    
    
    # the rows that were not returned appears as NAN values from teh merge
    # while you can overwrite the column values as "No", i will create the 
    # a new column so that you can see the before and after, and show you how 
    # you can create new columns as well
    
    orders_dataset['Return Status'] = orders_dataset['Returned'].fillna("No")
    
    # the next formula is to calculate the numbers of days to complete an order
    # we can just do normal algrbra on the columns just like you do in excel
    # the only different is that formulas are cell wise in excel, C2= A2+B2
    # in python you do newcolumn = column A + column B
    
    #when you are doing date time calculations, you might have to do some 
    # adjustment to make sure that the data is datetime calulcations
    # there are plenty of examples you can find for that online
    
    orders_dataset['Time to Complete Order'] =  orders_dataset['Ship Date'] - orders_dataset['Order Date']
    
    #---- Pivot tables
    
    #now we recreate the charts and visuals in the report
    
    #we are going to first do the pivot tables for time to completion
    # we will do it twice - once by region, then product type
    
    order_completeion_region_pivot = pd.pivot_table(orders_dataset, 
                                                    values='Time to Complete Order',
                                                    index='Region',
                                                    columns='Ship Mode',
                                                    aggfunc='mean',
                                                    margins=True, margins_name= 'Grand Total')
    
    
    # i will create a helper table that will help double check that the work 
    # that we are doing here actually is the same
    
    order_completeion_region_pivot_count = pd.pivot_table(orders_dataset, 
                                                    values='Time to Complete Order',
                                                    index='Region',
                                                    columns='Ship Mode',
                                                    aggfunc='count',
                                                    margins=True, margins_name= 'Grand Total')
    
    
    # we see that the time is formated as days hours minutes- we want to see this 
    # as fractions of a day instead
    
    # this is an example of the datetime math and adjustment that i warned you about
    # in excel this is done automatically, however we hace to be a bit more hands on
    # with python
    
    # this is where internet helps, and one solution is to  take the time to complete
    # and divide it by the total number of seconds in a day
    
    # I will create a copy of the df, this is not best practice as you are 
    # usually creating clutter and makes code management difficutl
    # however i do this here just to show an example of the calcualtion
    order_completion_region_pivot_temp = order_completeion_region_pivot.copy(deep=True)
    order_completion_region_pivot_temp['First Class'] = (order_completion_region_pivot_temp['First Class']/np.timedelta64(1,'D')).astype(float)
    
    # this solution had a different examples and names, however the solution
    # was similar enough that i only needed to change names and try the code
    
    #now that i have a line of code that does work formatting the data, i want to apply this 
    # to all the columns
    # of course i can copy paste everything and there is nothing wrong with that
    # however that defeates the purpose of coding we code to avoid copy and pasting repeatedable tasks
    # this can also have problems as you can forget to change names, or there could be more shipping methods adding in the future
    
    #python has a  much better solution, we can "loop" through a dataframe
    # and this allows us to apply the same line of code on multiple columns 
    # without copy and pasting the same repeatedly
    
    # not only does this allow us to avoid copy and pasting the same code over and over,
    # this makes the script easier to read so if there are mistakes, its easier to debug
    # and makes less mistakes from the programmer
    
    #we do this by creating a place holder for the column name with a variable
    
    for column in order_completeion_region_pivot.columns:
        print(column)
        order_completeion_region_pivot[column] = (order_completeion_region_pivot[column]/np.timedelta64(1,'D')).astype(float)
        
    #this looks great but lets just round the demicals out to make it nicer
    order_completeion_region_pivot = order_completeion_region_pivot.round(2)
    
    
    
    # the first part of the report is complete! 
    # that took a few steps, but now we can easily redo it for the summary by product
    # the only material change here is that we want to change the index param for the pivot table
    order_completeion_product_pivot =  pd.pivot_table(orders_dataset, 
                                                    values='Time to Complete Order',
                                                    index="Sub-Category",
                                                    columns='Ship Mode',
                                                    aggfunc='mean',
                                                    margins=True, margins_name= 'Grand Total')
    
    
    for column in order_completeion_product_pivot.columns:
        # print(column)
        order_completeion_product_pivot[column] = (order_completeion_product_pivot[column]/np.timedelta64(1,'D')).astype(float)
        
    #this looks great but lets just round the demicals out to make it nicer
    order_completeion_product_pivot = order_completeion_product_pivot.round(2)
    
    
    #the last pivot table that we want is return rate by region
    # we need to count the Yes and No by region
    
    # we previously used 'count' as a sanity check, but we can also use it to calculate the rate of return
    order_RR_product_pivot_count =  pd.pivot_table(orders_dataset, 
                                                    values='Order ID',
                                                    index="Region",
                                                    columns='Return Status',
                                                    aggfunc='count',
                                                    )
    #we dont do the margins params this time, as that just gives us totals,
    # and we dont need that here to calculate the return rate
    
    #in excel when we would do this, it would be by 'summarizing' the 
    # values as a percetn of row total
    
    # there are number of solutions to do this, this is the solution that
    # can be broken down into parts that are easier to understand
    
    order_RR_product_pivot_count.div(order_RR_product_pivot_count.sum(axis=1),axis=0)*100
    
    # this code takes advantage of a concept called chaining commands
    # not everything in python can be chained, but you can break things into parts
    # it is best to start from the inner most part and work outwards
    
    #inner most part is
    row_totals = order_RR_product_pivot_count.sum(axis=1)
    #we total by row, we know that we are summing by row, with the axis=1 
    # if axis=0, then the total is by column
    
    order_RR_product_pivot_count.sum(axis=0)

    #the outer part is 
    order_RR_product_pivot_count_dec = order_RR_product_pivot_count.div(row_totals,axis=0)
    #with the axis=0, the division is by column
    
    
    # now everything looks okay, but we want the summary in percent, not decimals
    # we need to finally multiply everything by 100
    
    order_RR_product_pivot_count_dec*100
    
    #so now we use this row, and save it down
    
    order_RR_product_pivot_count= order_RR_product_pivot_count.div(order_RR_product_pivot_count.sum(axis=1),axis=0)*100

    #as stated before - there are a number of solutions that could have done this, but
    # when i searched, this was the first soln that was paired with a clear explaination
    # that explains all the steps and how it achieves the final values

    #we are going to round the data again to make it easier to read    
    order_RR_product_pivot_count=order_RR_product_pivot_count.round(2)



    #---- Charting
    # remaing part of excel reporting the line chart
    
    #chart is only looing at 2017 data, lets filter the data out first
    orders_dataset_2017 = orders_dataset[orders_dataset['Order Date']>=pd.to_datetime('1/1/2017')]

    #when looking at the pivot chart fields in excel, we see that in axis categories
    # order data is aggregated at the month level
    # one way of doing this is creating a new column that just has month for each other
    
    #this is also a solution that was found online
    
    #strftime is short for 'string format time', in other words, 
    # we can format the order date and select the month
    orders_dataset_2017['Month'] = orders_dataset_2017['Order Date'].dt.strftime('%m')
    orders_dataset_2017['Month Name'] = orders_dataset_2017['Order Date'].dt.strftime('%b')

    #both of these trigger a SettingWithCopyWarning
    #these warning are important, and should be adddressed ASAP,
    #i use this solution because it makes sense, however this wont always work as upgrades take place
    #you see that the warning also directs you to a clear explaination and solutions that 
    # will avoid future problems
    
    #now that we have a nice column to aggregate out pivot table to create out chart on
    # we should just do a basic sanity check first
    
    
    profit_and_sales_by_region_count = pd.pivot_table(orders_dataset_2017,
                                                      index='Month',
                                                      columns='Region',
                                                      values='Profit',
                                                      aggfunc='count',
                                                      margins=True
                                                      )
    
    #all counts tie out, lets create the pivot table that shows profit and sales totals
    
    profit_and_sales_by_region = pd.pivot_table(orders_dataset_2017,
                                                      index=['Month','Month Name'],
                                                      columns='Region',
                                                      values=['Profit','Sales'],
                                                      aggfunc='sum',
                                                      # margins=True
                                                      )
    
    # the last step here is to convert the pivot table into a line chart and
    # save that down as a image
    
    #creating plots is very much the same charting by hand
    plt.figure(figsize=(100,50)) #this is settting the paper
    profit_and_sales_by_region.plot(kind='line') #drawing the line plot
    plt.title('Total Sales and profit all Regions 2023') #title
    plt.xlabel('') #clean up the x axis label to be blank
    plt.savefig(fr'{data_folder}\Profit and Sales.png')
    plt.show()
    
    
    #---- Create and send out email report
    
    #we have all the parts of the email report, lets actually recreate the email
    #this something that doesn follow from the out understanding of MS products
    
    #found the code creating and sending emails from online resources
    #this part can very as diffferent firm has different security config, so you need to be a bit careful
    #if using MS classic outlook, should be fine, its always worth double checking
    
    # first we "open" outlook
    outlook= win32.Dispatch('outlook.application')
    
    #now open the new email 
    mail = outlook.CreateItem(0)    
    
    #now we just add to the email similarly like we would if you were doing manually
    
    mail.To = "Shiv Mack @gmail.com; Super Store Manager@gmail.com"
    
    mail.Subject = "Super Store Market Report"
    
    mail.Attachments.Add(fr'{data_folder}\Profit and Sales.png')
    
    #create emails you need to use HTML formatting, most of this is pretty simple
    #<br> is line break

    mail.HtmlBody = 'Hello Manager, <br> <br> I have our updated analysis on Super Market. <br>'
    
    #now i want to add the tables to the email body
    #i can addd to the body using the +=
    # this is the same as doing mail.HtmlBody = mail.HtmlBody + "new stuff"
    
    mail.HtmlBody += "I have attached some relevant performance metrics and charts. <br> <br>"
    
    #i add the first pivot table to the body
    
    mail.HtmlBody += "Here is the return rate by region <br> " + order_RR_product_pivot_count.to_html() + "<br> <br>"
    
    #in the email we give the overall average o the time that it takes to complete an order, 
    # we can get the value out of the pivot table (there are many solutions online)
    # I will just find the overall average of the column and vonvert it to a number of days
    
    
    # I will use the same time rounding code that we used before
    
    average_completion_time = orders_dataset['Time to Complete Order'].mean() 
    average_completion_time =round((average_completion_time /np.timedelta64(1,'D')))
    
    mail.HtmlBody += f"We see the average time to complete an order is about {average_completion_time} days. "
    
    mail.HtmlBody += '<br>' + order_completeion_region_pivot.to_html() 
    
    mail.HtmlBody += '<br>' + order_completeion_product_pivot.to_html() 
    mail.HtmlBody += '<br> -ShivMakh'
    
    mail.Display()
    # mail.Send()
    
    
    
    
    """
    
    We are done!
    
    By Understand what the end goal is, and having an exisitng format that works we can literally take each step and translate that into pytho
    
    now all oyu have to do is change hte data set and you get an updated report
    
    building the report up to this point took some time, but now repeated times is going to be 
    less difficutl
    
    we have some very clear tangible results
    
    
    now again - this is just a proof of concept. there are improvement that can be done from a technical and business standpoint
    
    but now that you have this, you can go back and work on parts and make improvements on both ends as you run into problems
    
    and you can also make adjustments so this report can be a report for each regional mananger for their specific stores and so on
    
    
    as the managers review these reports, they might want different analysis and you can build on what you have here and create addtional reproducable reports with further insights too
    
    """
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    



