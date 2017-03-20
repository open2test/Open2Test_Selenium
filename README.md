# Open2Test_Selenium 

Open2Test_Selenium is Keyword driven Framework - supporting Continuous Integration using Gradle and Jenkins

### Synopsis: 

Open2Test framework is open source framework developed by Open source team of NTTDATA INC.
  This is framework developed for automation testing using Selenium. Aim of Open2Test framework is to provide easy way to approach automation testing. Keyword based solution formed on the premise that makes any application to be described using short text description or keywords.
  
 ### Salient Feature of Keyword Driven Testing

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

![Architectre](O2T_Architecture.png "Architecture Of Open2Test Framework ")

#### Work Flow Diagram of Open2Test Framework

![WorkFlow](O2T_WorkFlow.png "Work Flow Of Open2Test Framework ")

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
![Folder](Folder_Structure.png "Folder Structure of Imported Project")

#### Sample Scripting:
Please refer sample scripting from downloaded zip. 
1.	Change the path of variable utilityFilePath from setting.java file to SeleniumUtility file of downloaded /cloned folder. 
    E.g.  
Public String utilityFilePath=
              *"D:\downloaded_project\open2Test\o2t\\Sample script\\SeleniumUtility.xlsx";*
              
2.	Change the path of Test Suite, Object repository and other variables from utility file as per 
                             local downloaded folder structure.

3.	Select the appropriate browser. And make sure the browser driver is matching /compatible with selected browser. 




