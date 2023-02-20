# Automating weekly reconciliation of supplier payments

## What was the task?
The task was to automate supplier statement/invoice reconciliation for accounts payable. Suppliers will send through statements to QBT (travel agency) at the end of an agreed trading period (in this case a week) which will contain invoices that QBT has not yet paid. These statements (airticket statements) must be compared with the data within the QBT database (BSP tickets) to see if they matched. The statements are collated within an Excel workbook which will have over 20,000 rows of data in each page. As such, there are always discrepancies which will need to be reported under exceptional items. Exceptional items will be investigated and placed under categories such as small variances, commissions and airticket fees. Some invoices that come late which will be labelled as timing differences. 

## The automation process
The first step in starting this process was to understand the task and perform it manually. From this, a more deeper understanding could be attained allowing me to view and solve different scenarios. Once I was satisfied with my level of ability to do the reconciliation process manually, I moved onto writing code to automate the manual process. The steps involved are as follows:
1. User will be prompted to give the week ending date as well as the BSP report date.
2. User will be prompted to enter the period of airtickets (sheet name within the airticket statements). Dataframes of BSP tickets as well as airticket statements are initiated and stored for easy access later.
3. Add the BSP and airticket details together, remove duplicate tickets based on ticket number, then add time stamp sources to a new column to help indicate timing differences.
4. Create a new reconciliation dataframe using existing data from the BSP and airticket dataframes
5. Split the reconciliation dataframe into the different categories of exceptional items. Note columns named small variance, commissions, airticket fees and current timing differences
6. The previous weeks' ongoing timing differences are combined with the current weeks' timing differences. For those invoices that match exactly (BSP ticket + airticket statement = 0), they are grouped together and categorised as timing difference matched. 
7. As per the company's policy for accounts payable to be paid is 2 weeks, the earliest ongoing timing differences tickets from 2 weeks ago, which are shown through the time sources column, will be processed via the same method as step 6. The leftover tickets from this time source will be split accordingly between small variances, commissions and airticket fees. Some special cases will be reviewed by management.
8. All previous datasets are written to an Excel workbook for ease of access when need be.
9. Once the ongoing timing differences from two weeks ago have been sorted, a coversheet is created showing the exceptional items which is kept for recording purposes. This coversheet contains the data for small variances, commissions, airticket fees and the now-ongoing timing differences which will be passed on to next week.

## Problems encountered and how I overcame them
As this process was quite time-consuming and had a lot of steps, creating the script would have a lot of pieces. As such, I split the task into achieveable parts, allowing myself to be less stressed and be able to start comfortably. Many times when my script would not work as intended, I spent much time googling and experimenting with smaller pieces of code. Although a lot of time was lost doing that, my goal was achieved and I grew my ability to spot and fix errors quickly. 

## Results and Achievements
Before automation, when other people and I were doing the reconciliation process manually, it took a person one day to process and investigate one weeks worth of statements. After writing and running the script, it took one hour to process the datasets, create a coversheet and investigate the exceptional items for each week. This was almost an increase of 8x in efficiency for this task. 
