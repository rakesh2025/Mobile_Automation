<?xml version="1.0" encoding="UTF-8"?>
<TestSuiteCollectionEntity>
   <description>Execute Both the test case and verify the data</description>
   <name>TSC_01_Brandfactory_Dataverify</name>
   <tag></tag>
   <executionMode>SEQUENTIAL</executionMode>
   <maxConcurrentInstances>8</maxConcurrentInstances>
   <testSuiteRunConfigurations>
      <TestSuiteRunConfiguration>
         <configuration>
            <groupName>Mobile</groupName>
            <profileName>default</profileName>
            <runConfigurationData>
               <entry>
                  <key>deviceName</key>
                  <value>Xiaomi M2003J15SC (Android 11)</value>
               </entry>
               <entry>
                  <key>deviceId</key>
                  <value>7bbbca200408</value>
               </entry>
            </runConfigurationData>
            <runConfigurationId>Android</runConfigurationId>
         </configuration>
         <runEnabled>true</runEnabled>
         <testSuiteEntity>Test Suites/TS_01_Brandfactory_OnMobile</testSuiteEntity>
      </TestSuiteRunConfiguration>
      <TestSuiteRunConfiguration>
         <configuration>
            <groupName>Web Desktop</groupName>
            <profileName>default</profileName>
            <runConfigurationId>Chrome</runConfigurationId>
         </configuration>
         <runEnabled>true</runEnabled>
         <testSuiteEntity>Test Suites/TS_02_Brandfactory_OnWeb</testSuiteEntity>
      </TestSuiteRunConfiguration>
   </testSuiteRunConfigurations>
</TestSuiteCollectionEntity>
