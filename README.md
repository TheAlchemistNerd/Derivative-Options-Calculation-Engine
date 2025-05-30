# Option Analysis Project

This project is a Java-based tool for analyzing stock option prices using historical stock data and the Black-Scholes model. It loads stock price data and risk-free rate data from an Excel file, calculates historical volatility, and then uses these parameters along with specified strike prices and time to expiration to estimate call and put option prices.

## Overview

The project performs the following key steps:

1.  **Data Loading:** Reads historical stock prices and Treasury bill rates from different sheets within an Excel file (`SPY_Data_2024.xlsx`).
2.  **Return Calculation:** Computes logarithmic daily returns from the historical stock prices.
3.  **Volatility Calculation:** Calculates the annualized historical volatility of the stock based on the daily returns.
4.  **Risk-Free Rate Calculation:** Determines the average risk-free rate from the Treasury bill data for specified periods.
5.  **Strike Price Definition:** Defines a set of strike prices relative to the initial spot price of the stock.
6.  **Black-Scholes Option Pricing:** Implements the Black-Scholes model to calculate theoretical prices for both call and put options.
7.  **Results Output:** Prints the calculated option prices for different strike prices and time periods.

## Files

* `OptionAnalysis.java`: Contains the main Java code for the project, including methods for data loading, calculation of returns, volatility, risk-free rate, strike prices, and Black-Scholes option pricing.
* `SPY_Data_2024.xlsx`: (Example) An Excel file containing historical SPY ETF prices in sheets named "Mar-Jun 2024" and "Jul-Oct 2024", and Treasury bill rates in a sheet named "TB3MS_2024". The stock price sheets are expected to have dates in the first column and closing prices in the second column. The Treasury bill sheet should have dates and the corresponding rates.
* `.gitignore`: Specifies intentionally untracked files that Git should ignore.

## Dependencies

This project uses the following external libraries:

* **Apache POI:** For reading data from Excel files (`org.apache.poi:poi`, `org.apache.poi:ooxml-schemas`, `org.apache.poi:poi-ooxml`).
* **Apache Commons Math:** For statistical calculations, specifically the normal distribution (`org.apache.commons:commons-math3`).
* **Apache Log4j 2 (API and Core):** For logging within the application (`org.apache.logging.log4j:log4j-api`, `org.apache.logging.log4j:log4j-core`).

These dependencies are managed using Maven, and their definitions can be found in the `pom.xml` file (if this is a Maven project).

## Setup and Usage

1.  **Prerequisites:**
    * Java Development Kit (JDK) installed.
    * Maven installed (if using Maven for dependency management).
    * An Excel file named `SPY_Data_2024.xlsx` (or your data file) in the project directory, with the specified sheet names and data format.

2.  **Installation (if using Maven):**
    * Navigate to the project directory in your terminal.
    * Run the command `mvn clean install` to download dependencies and build the project.

3.  **Running the Application:**
    * **Using IDE:** Open the `OptionAnalysis.java` file in your IDE and run the `main` method.
    * **Using Maven:** Navigate to the project directory in your terminal and run the command `mvn exec:java -D"exec.mainClass"="OptionAnalysis"`. You might need to adjust the classpath if you are not using Maven. For example:
        ```bash
        java -classpath ".;path/to/poi-*.jar;path/to/poi-ooxml-*.jar;path/to/commons-math3-*.jar;path/to/log4j-api-*.jar;path/to/log4j-core-*.jar" OptionAnalysis
        ```
        (Replace the `path/to/...` with the actual paths to your JAR files.)

## Output

The application will print the calculated call and put option prices for the specified strike prices for the "Mar-Jun 2024" and "Jul-Oct 2024" periods to the console. The output will be in the format: