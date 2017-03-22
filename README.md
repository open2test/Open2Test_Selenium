# Open2Test_Selenium     ![o2t_logo](https://cloud.githubusercontent.com/assets/25658523/24182607/f9231ce6-0ee8-11e7-821d-cbfe20dda15b.jpg)

Open2Test_Selenium is Keyword driven Framework - supporting Continuous Integration using Gradle and Jenkins

### Synopsis: 

Open2Test framework is open source framework developed by Open source team of NTTDATA INC.
  This is framework developed for automation testing using Selenium. Aim of Open2Test framework is to provide easy way to approach automation testing. Keyword based solution formed on the premise that makes any application to be described using short text description or keywords.
  
  ### Index
  
 
  [Features](<####Salient-Features-Of-Keyword-Driven-Testing>)
  
  [Architecture](####Architecture-Of-Open2Test-Framework)
  
  [Workflow](####Work-Flow-Diagram-of-Open2Test-Framework)
  
  [Installation](###Installation–Instructions-for-Automation-Beginner)
  
  
  
#### Salient Features of Keyword Driven Testing
 
#### Natural language: - 
                       Keywords can be defined in natural language meaning that, test cases can be 
                       written with more or less detail, depending on the project's needs.

#### Not require technical understanding:-
                         Testers working at the business level do not require technical understanding 
                         of the test automation framework to be able to create and edit test cases

#### Maintenance:-
                      Sensitivity to changes (which can create the need for maintenance effort) is reduced.
                      Maintenance of test cases for business reasons will not affect the implementation 
                      of the Keyword scripts, as long as the set of keywords and the semantics of the
                      keywords are not changed.

#### Portability:-
                       Portability of test suites is easier to achieve, (e.g. if a similar system with 
                       almost the same business cases has to be tested then many of the keywords can be reused).

#### Early Start:-
                    Automated functional tests can be implemented before the test item actually exists,
                    either by sing existing keyword libraries with their corresponding automation scripts,
                    or by defining new keywords and adding the automation scripts later as the test 
                    interface is defined.

#### Efforts:- 
                     A limited set of keywords implies a limited effort for implementing test automation, 
                     (e.g. usually, one automated keyword script for each keyword will be sufficient).


####            Please refer Appendix A – for further Keyword driven testing information 

####  Architecture Of Open2Test Framework 

![o2t_architecture](https://cloud.githubusercontent.com/assets/25658523/24182574/c84894b6-0ee8-11e7-98e5-8fdb4a1f91f9.png)

#### Work Flow Diagram of Open2Test Framework

![o2t_workflow](https://cloud.githubusercontent.com/assets/25658523/24182611/fef3e484-0ee8-11e7-88c9-e1515e30a53f.png)

### Installation – Instructions for Automation Beginner  

|Tools/Software |	Version |	Verified Version |	Comments |
|---------|:--------|:-----------------|----------:|
|Eclipse	|Luna or above|	Luna	|Editor|
Java	|1.7 Or above|	1.8.52|	O2T Framework Developed Lang|
Gradle| 	3.2.1|	3.2.1|	Build Tool| 
Selenium *[with all  jar from Libs folder]* |	2.53.1 and above|	2.53.1 and 3.1|	Automation Tool|
Apache POI jar library|	3.13|	3.13|	For Excel interaction|
Ashot jar library| 	1.5.3|	1.5.3|	To take Screen shot in Chrome|
Junit| 	4.12|	4.12|	Annotations| 
Jenkins *[Jenkins is not required, in case of local configuration.]*|	2.3X|	2.3X	|open-source continuous integration software tool|
Jenkins   Plug ins *1.	Gradle Plugin 2.	Subversion release manager plugins.*|   |   |	Supporting Plug ins for Continuous integration| 
Subversion [SVN ] Repository|	1.6.12|	1.6.12|	Code repository|
Open2Test|   |   |			Keyword driven framework|


### Configuration:
#### Open2Test:
Download / Clone open2Test Selenium Framework of Selenium WebDriver from Git-hub.
#### Java & Eclipse:   
   First install Java   
   
   Set Java_Home and Path in environment variable 	
   
   
   e.g.  Java_Home = C:\Program Files\Java\jdk1.8.51
   
   
                        Path =; %JAVA_HOME%\bin;

 Unzip Eclipse folder and launch eclipse with Eclipse exe. 
 

#### Gradle: 

Traverse to URL: <https://gradle.org/gradle-download/>
Then click on Complete Distribution link 
(Complete distribution (binaries, sources and offline documentation))
Gradle-3.2.1-all.zip file will get download. [Current version is 3.3]

 
In Eclipse --->  Help ---> Marketplace ---> Search and Install plug in called *“Buildship Gradle Integration2.0”*

### How to Use
                     Import the downloaded Open2Test Project 
          
#### Importing Open2Test Gradle Project

1. Open Eclipse 

2. File menu ==> Import==> Gradle Project

3. Click on Next    Welcome pan of Gradle project import wizard will get open 

4. Click on Next   Import Gradle Project pan/page will get open

5. Enter /browse "Project Root Directory" as path to your downloaded open2Test 

       Project.     e.g. D:\downloaded_project\open2Test
       
6. Click on Next      Import options pan/page will get open

7.  Select radio button "Local Installation directory"          
                                              browse to downloaded    Gradle zip
          e.g. "D:\downloaded_project\Gradle\gradle-3.2.1-all\gradle-3.2.1"
          

8. Click on Next Button - Import preview pan will get open 

                           Review the entered details like Root Directory , Gradle version etc. 

9. Click on Finish Button


If everything goes well, Gradle project will get import with following folder structure 

![folder_structure](https://cloud.githubusercontent.com/assets/25658523/24182616/07bac4e8-0ee9-11e7-8f67-a71db53c7983.png)


#### Sample Scripting:
Please refer sample scripting from downloaded zip. 
1.	Change the path of variable utilityFilePath from setting.java file to SeleniumUtility file of downloaded /cloned folder. 
    E.g.  
Public String utilityFilePath=
              *"D:\downloaded_project\open2Test\o2t\\Sample script\\SeleniumUtility.xlsx";*
              
2.	Change the path of Test Suite, Object repository and other variables from utility file as per 
                             local downloaded folder structure.

3.	Select the appropriate browser. And make sure the browser driver is matching /compatible with selected browser. 

#### Executing Gradle Project on Local Machine: 
 Traverse to Root folder of the project on Command prompt.
1. Use the command “gradle build” on Command prompt for executng the build . . . 	[Gradle Executin]
2. Right Click on SeleniumWD.java file and select Run as Gradle Tests.     ……[Gradle Execution]
3. Right Click on SeleniumWD.java file and select Run as Junit Tests.     ……[Java Execution]

This will run the sample script file available in downloaded zip folder. 

User can change the script and Object repository as per project requirement and continue to use. 

#### Jenkins –                       *[Optional in case of Local configuration]*
##### Download

  * Download the Jenkins from <https://jenkins.io/>      … Zip file will get download.*
  * Download Jenkins.war  <https://updates.jenkins-ci.org/download/war/> *    
  
*Unzip the Jenkins folder and paste the downloaded war (“Jenkins.war”) File to Jenkins folder.*


#### Subversion [SVN]:
Download and configure SVN.
<https://subversion.apache.org/download.cgi>

#### Commit Project folder to SVN Repository 

*1. Go to workspace directory of current project.*

*2. Copy Project folder and past it to SVN repository.*

*3. Commit and update the SVN.*

#### Jenkins Project Configuration 

Traverse to Jenkins folder on command prompt
Execute the command    “ Java -jar jenkins.war” 

Once Jenkins is fully up and running, open URL in browser 
<http://localhost:8080>
Jenkins log in screen will get displayed. 

By default – Jenkins username is Admin and password can be get it from initialAdminPassword.txt file which is available in secrets folder of Jenkins
Jenkins Dashboard will get open 

Add Plugins : 
Add following plugins by traversing on Manage Jenkins ---> Manage PlugIns
--->Gradle Plugin
--->Subversion release manager plugins.

##### Create new Item in Jenkins. 

*	Click on New Item -to create the Build
-	Select the Project Type - Free Style
-	Fill the general Tab - Enter Description
-	In Source code management - Select Subversion
-	Enter Required Details Repository URL and Credentials
-	Click on Add build step  - Select Invoke Gradle script as it is Gradle project

                  *(Build Tab will get open with the Gradle details - No change required in this)*
-	Click on Save Button
-	Demo build page will get enable - where Build now link is visible
-	Click on Build Now link 
-	Check the status on Console

### How Open2Test works
Let’s understand, how Open2Test for selenium framework works
   #### Higher Level Overview

![higher_overview](https://cloud.githubusercontent.com/assets/25658523/24182621/0b05d2dc-0ee9-11e7-8c56-350132c0563d.png)

   #### Read Utility File
   
 ![read_util](https://cloud.githubusercontent.com/assets/25658523/24182627/1980b5ac-0ee9-11e7-8878-2f9e2fe46d7a.png)
   
   #### Script Execution
   
   ![script_execute](https://cloud.githubusercontent.com/assets/25658523/24182650/277a74cc-0ee9-11e7-848e-3362827b3686.png)
   
   
   #### Test Suite to Test Script
![testsuite_testscript](https://cloud.githubusercontent.com/assets/25658523/24182668/3430d4ae-0ee9-11e7-8657-911301c2405c.png)
   
   #### Test Scripts Reads Test Data
   
  ![testscript_testdata](https://cloud.githubusercontent.com/assets/25658523/24182662/2dd0a742-0ee9-11e7-8d4c-d824640668d9.png)
   
   
### Introduction to O2T Keywords
#### Keyword: "launchapp"
#### Brief Description:
            Launches the given URL
#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|launchapp|https://www.google.com |   |   |Launching the Website|

In above table – These are the Excel columns with Keyword launchapp given in second {B} column  
  and in Object Details column  [C] has the URL to launch.
  

#### Keyword: "importdata"

#### Brief Description: Imports the Data from the given path (Data Sheet path)

#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|importdata|E:\AAVmwareDemo\ |   |   |Imports  the Data from the given path (Data Sheet path)|


#### Keyword: "screencaptureoption"
#### Brief Description: Captures the screen shot for all the Perform/Storevalue/Check test steps, if indicated in the Object column
#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|screencaptureoption|Perform;>>;Storevalue;Check |   |   |Captures the screen shot for all the Perform test steps,if indicated in the Object column|
                                                
##### EXAMPLE 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|screencaptureoption|perform |   |   |   |


#### Keyword: "perform"
#### Brief Description: Captures the screen shot for all the Perform/Storevalue/Check test steps, if indicated in the Object column
#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|perform|  <TextBox/ CheckBox/ RadioButton/ Button/ Link/ ComboBox/ TextArea/ Image/ Frame/ iFrame/ Table/element>;<Object Name>|click, altclick;enter; hover; Select; Set; Check etc.|   |Clicks/Hovers/Select/Set the Selected/located object|

##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|perform|Button;btn_Submit | Click  |   |  Clicks on Button which locator [ID ,Xpath, etc.] is stored in OR and name is given as btn_Submit by user. |

                                                
#### Keyword: "check"
#### Brief Description: Returns whether object is Visible or not. If so True, else False; Returns whether object is enabled or not. If so True, else False; Checks the Text property of the given object.checks the displayed text of the link.Validate the selected item from the combobox.Data table variables/ Environment variables can be used here for value to compare 

#### Syntax 

|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|Check|  <TextBox/ CheckBox/ RadioButton/ Button/ Link/ ComboBox/ TextArea/ Image/ Table/ListBox/element>;<Object Name>|Visible:<True/False> enabled:<True/False> text:<Text to compare> linktext:<Text to compare> value:<value to be compare> checked:<ON/OFF>.|   |Returns whether object is Visible /Enable/ or not. Compares the text, Checks the link text etc.|

##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|check|Button;btn_Submit | enabled:<True/False>  |   |  Checks if btn_submit is enabled or not |

#### Keyword: "storevalue"
#### Brief Description: Assigns the display status of the element in a variable. 
#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|storevalue|  <TextBox/ CheckBox/ RadioButton/ Button/ Link/ ComboBox/ TextArea/ Image/ Table/ListBox/element>;<Object Name>|Visible:<Variable Name> enabled:<Variable Name> value:<VariableName> linktext:<VariableName>|   |Assigns the text value / display status / object is enabled or not/ Assigns selected / Stores the linktextitem  of the element to a variable|

##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|storevalue|Button;btn_Submit | enabled:variable1  |   |  It will assign the True value to variable variable1 if button is enabled else it will assign false. |

#### Keyword: "loop" . . . "Endloop"
#### Brief Description: Executes the statements inside the loop for <count> iterations. All the statements in the between  loop ..endloop will get execute till given count.
#### Syntax 
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|loop|  count|   |It will start iterating from this row|
| r	|endloop|    |    |It will stop iterating on this row|



##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|loop|3 |    |    |  Iteration will start - and it will iterate for 3 times |
| r |perform|	Button;btn_Submit|	click|    |Clicks the button|
| r	|storevalue|Button;btn_Submit | enabled:variable1  |    |  It will assign the True value to variable variable1 if button is enabled else it will assign false. |
| r	|endloop|    |    |    |  stops the loop  |

#### Keyword: "context"
#### Brief Description: Use to switch the context to the frame / parent Frame / iframe/child frame/page/paagetitle. 
 Once you switch the context to the frame, you will be able to access and work on the elements inside the frame.
#### Syntax
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|context| <frame/iframe;frame object name> <browser;WebBrowser> <frame/iframe;browser;WebBrowser> <Page;pagetitle browser;WebBrowser>|  page;<pagetitle>::<frame/iframe>;<frame object name> page;<pagetitle>::<frame/iframe>; page;<pagetitle>|    |Use to switch the context to the frame / parent Frame / iframe/child frame/page/paagetitle|


##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|context|Frame;UserDetails | enabled:Frame1  |   |  Control will switch to the Frame1. |
| r	|context|Browser;Individual Health Plans | page:Individual Health Plans |   |  Control will switch to page of name “Individual Health Plans |

#### Keyword: "upload"
##### Brief Description:  Upload the file. Precondition: A file upload dialog is Open If Cancel/Close action is given then closes the upload dialogue 
#### Syntax 

|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|upload| File path  in C column of Test Script|     |    |Upload the file . Precondition: A file upload dialog is Open|
| r	|upload| File path  in C column of Test Script|  closeupload/ cancel upload   |    |Close file upload dialog. Precondition: A file upload dialog is Open|

##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|upload|C:\Users\076533\Desktop\OLDDESK\dcxccz.png|    |    |  Upload the file  |
| r	|upload|C:\Users\076533\Desktop\OLDDESK\dcxccz.png | cancelupload |   |  Close file upload dialog |

#### Keyword: perform    *(Calendar control)*
#### Brief Description:  It will set date in calendar control, with condition that object type is Calendar and object name should be prefixed by cal_
#### Syntax 
Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|perform  | calendar:cal_Calendar1|  setdate:<Date in mm-dd-yyyy format only>|    | It will set date in calendar control|

*Object type is calendar and object name will start with cal_ on which this keyword can be applicable. Column Action is having Action as Setdate which accepts date given in mm-dd-yyyy format only*


##### EXAMPLE
|Run|	Keyword|ObjectDetails |Action|Action2|Comments|
|---|:-------|:-------------|:-----|:------|-------:|
| r	|perform  |calendar;cal_calendar1|setdate:dt_12-21-2017 |    |  It will set the date as 21 Dec 2017 in selected calendar. |


### Appendix A

#### Layers in Keyword-Driven Testing:
Keyword-Driven Testing can be organized by using one or more layers. Typical layers are

 #### End user domain layer 
 #### Test interface layer.

While many implementations of Keyword-Driven Testing will comprise two or three abstraction layers, 
in some cases it may be necessary to structure keywords in more layers.

#### Abstract layer:
The topmost layer is the most abstract layer which is generally aligned with the wording of the application's users. 
In practice, the topmost layer is usually the domain layer. 
However, in some situations the domain layer may not be required, and another, more abstract layer is used 
(e.g. if the test cases are supposed to span several different end user domains, a Meta domain layer can be introduced).

#### Detailed layer:
The lowest layer is the most detailed layer. It is commonly aligned with the names of test interface elements   
in practice, this layer is usually, but not always, the test interface layer (e.g. asSometimes a test interface layer 
is not required, or for specific reasons, even more detailed layers may be used).  
    Most Keyword-Driven Test systems will have more than one layer due to factors such as having understandable
    keyword test cases, maintainability and division of work relying on a multi-layer structure. If only one layer 
    is implemented, it will commonly be either at a low level, which affects the readability of the keyword test cases,
    or at a high level, which can result in more keyword execution code.
    
 ### Types of keywords
 
#### Simple keywords
       
Simple keywords, which are often used at the test interface layer (e.g. "MenuSelect" or "PressButton"), can be the
connection between the test execution tool and higher level keywords at an intermediate layer or domain layer.
    Using only keywords at the test interface layer can be sufficient for the definition of test cases and their execution. 
    Exclusive use of simple keywords will lead to test cases with many actions. Depending on the test item, keywords at the test
    interface layer could need to interact with different systems such as databases, the system registry or SOA-Messages.
           In a similar way, the automation framework will support access to the test interface or other interfaces on
           which the keyword operates (e.g. mouse, keyboard, and touch screen).
           
                 
#### Composite keywords
        
Simple keywords are sufficient to compose and execute test cases but are often insufficient to reflect 
functional features. Composite keywords are keywords composed from other keywords. This means that keywords 
can be organized in different layers (see 5.2). For composite keywords, composite parameters (e.g. a data structure) 
can be required. It is often useful to use business-level keywords, such as “login user“. This keyword may be composed
of a sequence of lower level keywords, such as “enter username“, “enter password“ and "push login-button“. 
For more complex business objects, such as large forms for the preparation of contracts, a keyword like 
“filloutContractformPage1” can be valuable.
   A composite keyword is a ‘package’ containing a sequence of other keywords. The set of parameters for a 
composite keyword can be the union of the set of parameters of the keywords that comprise the composite keyword;
sometimes however, the implementer of a composite keyword may choose to ‘hide’ one or more parameters by 
assigning it a literal value within the composite.
       
 #### Navigation/interaction (input) and verification (output)

 *Keywords may be classified into at least two categories:*
 
#### Navigation steps (i.e. input to the test item) and
#### Verification steps (i.e. output from test item).

Most keywords belong to the first category, (i.e. the navigation steps) because most actions are needed to prepare
the test item or perform certain actions on it which will lead to a result. Navigation steps usually are steps 
that do not verify and log the test result.


The result is then checked by one or more other actions i.e. the verification steps.

The verification steps are related to the result of the test case. For example, if the condition of a verification
step is not met, then the test result will be set to "failed".

It may be useful to allow navigation steps to be used for verification.

 #### Keyword Driven Framework
   
  ![keyword_driven](https://cloud.githubusercontent.com/assets/25658523/24182623/11bde722-0ee9-11e7-81cb-d05f58cfa516.png)


#### Keywords and Data
Keyword-Driven Testing can be enhanced if keywords are associated with data. To allow an association with data,
in many cases keywords will need to have parameters which may be fixed, or list driven.
       Most keywords will need to have at least one parameter to specify the object they apply to. 
       Some will need another parameter to specify input, (e.g., true/false, a string to type, an 
       option to select in a combo box). This input will generally depend on the type of control and the type of action.
       
       







