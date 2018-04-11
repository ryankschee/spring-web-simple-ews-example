# spring-web-simple-ews-example
Simple example to show how to access MS Exchange via ews in Spring Boot. 

# How to Run
1. Start application 
```
mvn spring-boot:run
```
2. Run curl command
```
curl --get http://localhost:8080/ews/v1/readEmails
```