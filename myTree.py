#!/usr/bin/env python

# http://www.cnblogs.com/lhj588/archive/2012/01/06/2314181.html
# using python to read xlsx file

# https://github.com/caesar0301/treelib
# http://treelib.readthedocs.io/en/latest/

# this script is used to convert the proj-1955 Test Scenarios excel sheet
# into the tree data structure for the following:
# CP, Feature, User Story, Test Scenarios
# 


import xlrd
import re
from treelib import Node, Tree

# open Excel file


def open_excel(file= 'sample.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
    #except:
        print str(e)
        #pass

#data = xlrd.open_workbook('sample.xlsx')

def main():
    # open Excel file 'sample.xlsx'
    data = open_excel(file='sample.xlsx')
    
    # open the sheet "1955_Allocation" which defines the relationship
    # between cp, feature, user-story
    table = data.sheet_by_name(u'1955_Allocation')

    # title_raw the raw non-formalized title list
    cp_raw = []
    cp_title = []
    f_raw = []
    f_title = []
    s_raw = []
    s_title = []
        
    
    cp_raw = table.col_values(1)[1:]
    cp_title = table.col_values(2)[1:]
    f_raw = table.col_values(3)[1:]
    f_title = table.col_values(4)[1:]
    s_raw = table.col_values(5)[1:]
    s_title = table.col_values(6)[1:]
    
    #print type(f_raw)
    #print str(f_raw[0])
    #print str(int(s_raw[0]))
    
    #print "The number of user story: ", (len(cp_raw))
    
    # Format the feature & user_story, to turn (float & int) into string.
    # also do the format for the titles: cp_title, f_title, s_title, to remove ' & " from it.
    
    for i in range(len(cp_raw)):
        
        s_raw[i] = str(int(s_raw[i]))
        #print 's' + s_raw[i] + ', ',
        
        f_raw[i] = str(int(f_raw[i]))
        #print 'f' + f_raw[i] + ', ',
        
        cp_title[i] = re.sub(r'\'|\"', '', cp_title[i])
        f_title[i] = re.sub(r'\'|\"', '', f_title[i])
        s_title[i] = re.sub(r'\'|\"', '', s_title[i])
        
    #cp_raw.pop(0)
    #print f_raw
    #print s_raw

    # define the tree
    
    tree = Tree()
    tree.create_node('1955 Routine Maintenance - SDP automation and integration', '1955')
    
    # create node for CPs, they're under 1955
    # tree.create_node('CP-6541 Work order lifecycle and integration with field service system', 'CP-6541', parent='1955')
    for i in range(len(cp_raw)):
        
        # Add CPs
        # tree.create_node('cp_title', 'cp', parent='1955')
        try:
            tree.create_node(cp_raw[i] + ' ' + cp_title[i], cp_raw[i], parent='1955')
        except:
            pass
        
        #"""
        # Add Features
        try:
            tree.create_node('f' + f_raw[i] + ' ' + f_title[i], f_raw[i], parent=cp_raw[i])
        except:
            pass

        # Add user Stories
        try:
            tree.create_node('s' + s_raw[i] + ' ' + s_title[i], s_raw[i], parent=f_raw[i])
        except:
            pass
        
       # """
                        
    #tree.show()
    
    # open the sheet "SIT Test Scenarios-FTTX"
    table2 = data.sheet_by_name(u'SIT Test Scenarios-FTTX')

    # define a dictionary for user_story_id and user_story_title
    #s_dict = {}
    #s_dict = dict(zip(s_raw, s_title))
    #print s_dict
    
    # init the variable list for scenario and user-story
    scena_raw = []
    us_raw = []
    
    
    #scena_raw = table2.col_values(11)[2:] # get the Test Scenario title from Column L
    scena_raw = table2.col_values(12)[2:] # get the Test Scenario title from Column M
    #scena_raw = table2.col_values(11)
    us_raw = table2.col_values(3)[2:]
    #us_raw = table2.col_values(3)
    #print scena_raw
    #print us_raw
    
    #print '\n\n\n'
    #print 'len of Scenario: ', len(scena_raw)
    #print 'len of user-story: ', len(us_raw)
    #print type(us_raw[0])
    #print type(us_raw[-1])
    
    # init the scenario and story list
    us_lst = []
    scena_lst = []
    
    # first remove the blank one u'' from Test Scenarios
    # print 'len of Scenario: ', len(scena_raw)
    i=0
    while i < len(scena_raw):
        if scena_raw[i] != u'':
            title = re.sub(r'\'|\"', '', scena_raw[i])
            
            # if user story id contains more than 1 story,
            # split it by '\n'
            if type(us_raw[i]) == unicode:
                #print us_raw[i]
                tempUS_lst=us_raw[i].split('\n')
                for US in tempUS_lst:
                    #print US
                    us_lst.append(str(int(US)))
                    scena_lst.append(title)   
            else:
                us_lst.append(str(int(us_raw[i])))
                scena_lst.append(title)
            
        i = i+1
    
    #print i 
    #print us_lst   
    #print scena_lst
    #print len(us_lst)
    #print len(scena_lst)
    
    # Add user Stories
    for i in range(len(scena_lst)):            
        try:
            tree.create_node(scena_lst[i], parent=us_lst[i])
            #pass
        except:
            pass
    
    
    tree.show()
    
    
    
    
    
    

if __name__ == "__main__":
    main()
