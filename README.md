# caretool
Tool for my mother in law and her colleagues to generate care time table.
The MS excel care schedule is used for billing purposes as well as calculating overtime hours, etc.

The spreadsheet used to be created manually every year. However it's time consuming as well as error-prone.
So this project came to life. 

## Execution
To generate the jar
```
./gradlew jar
```

Execute the jar
```
java -jar build/libs/caretool.jar --names=Foo,Bar --year=2019

```

## Versions
### 0.0.1
- usage of jollyday library to retrieve the hollydays
- cmd interface to create the timetable

## upcoming features
- UI to easyly generate the timetable


