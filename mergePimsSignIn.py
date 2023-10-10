# written by Werner Goveya as a hobby project
# released under GPL v3
#DISCLAIMER: released as is, no warranty, liability, implicit or explicit.

# for privacy, no data files are included.

# this notebook allows you to merge a signin sheet and pims list
# useful for extracing latest email addresses and other columns unique to one of the worksheets
# pre-reqs: for simplicity, rename the signin sheet to be Sign.xlsx and pims to be PIMS.xlsx
# place them in the same folder as this script

# variables used:
# signInImp, signInMerge, pimsImp, pimsMerge

# pre-reqs
import pandas as pd
import numpy as np
from datetime import datetime

signInImp = pd.read_excel("Sign.xlsx", 
                          header=2, 
                          parse_dates = True, 
                          index_col = 0
                         )
signInImp.head(3)

signInImp.columns = signInImp.columns.str.replace(' ', '')
#signInImp

cols = signInImp.columns
cols

#delete non required columns
signInImp = (signInImp
             [cols]
             .drop(columns = ["Signature",
                             "Unnamed:10",
                              "Unnamed:6"
                             ])
            )

signInImp.head(3)

signInImp.shape

# TEST THE DATA
# enter the memberId and last name to get information about the person

memberId = 123123123
signInImp.loc[[memberId]]

lastName = "Goveya"
signInImp.loc[signInImp["LastName"] == lastName]

# Export the data
signInImp.to_csv("signInImp.csv")

cols = signInImp.columns
signInMerge = (signInImp
             [cols]
             .drop(columns = ["LastName",
                             "FirstName",
                              "BU",
                              "Member?",
                              "PAC",
                              "RegToVote"
                             ])
            )

signInMerge.head(3)


# TEST THE DATA
# enter the memberId and last name to get information about the person

memberId = 123123123
signInMerge.loc[[memberId]]


lastName = "Goveya"
signInMerge.loc[signInImp["LastName"] == lastName]



# import pims worksheet

pimsImp = pd.read_excel("PIMS.xlsx", header = 1, index_col= 0)
pimsImp.columns = pimsImp.columns.str.replace(" ","")
pimsImp.head(2)

pimsImp.index.names = ['MemberId']
pimsImp.head(2)

# TEST THE DATA
# enter the memberId and last name to get information about the person

memberId = 123123123
signInMerge.loc[[memberId]]

pimsImp.shape

cols = pimsImp.columns
cols

pimsImp = (pimsImp
           .drop(columns = ["Campus","A/P","Agency","Classification","PAC",])
          )


pimsImp.shape

pimsImp.head(2)

pimsImp.columns

# filter to the factors you need, example chapter, current members

# you can update this query to filter on additional columns
chapter = 307
#hireDate = "2023-07-01"

pimsMerge = pimsImp[
    (pimsImp["Chapter"] == chapter) 
#    & (pimsImp["Member"] == "Yes") 
#    & (pimsImp["HireDate"] > hireDate)
]
pimsMerge.shape


# export the file
exportFileName = "pimsMerge.csv"
pimsMerge.to_csv(exportFileName)


#test the data
lastName = "Goveya"
pimsMerge.loc[pimsMerge["LastName"] == lastName]




# merge the two files, pims and signin 



mergePimsSignIn = pd.concat([signInMerge, pimsMerge], axis = 1)
mergePimsSignIn.head(2)

mergePimsSignIn.shape


# TEST THE DATA
# enter the memberId and last name to get information about the person

memberId = 123123123
mergePimsSignIn.loc[[memberId]]

# export to csv the merged file
mergePimsSignin.to_csv("mergePimsSign.csv")



# you can update this query to filter on additional columns or use the merged csv which you just exported.

# LastName = "Goveya"
hireDate = "2023-07-01"

myQuery = mergePimsSignIn[
    (mergePimsSignIn["HireDate"] >= hireDate) 
]
myQuery


