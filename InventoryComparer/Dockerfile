FROM eclipse-temurin:17-jdk-alpine
WORKDIR /app
# If using Maven, JAR is in target/
COPY target/*.jar app.jar
EXPOSE 8080
CMD ["java", "-jar", "app.jar"]
