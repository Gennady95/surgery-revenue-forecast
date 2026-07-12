# Прогноз продаж по планируемым операциям
Скрипт анализирует планы лечения и запланированные операции, рассчитывает минимальную и максимальную ожидаемую выручку по каждому пациенту и формирует Excel-отчёт. Используется для прогнозирования будущих продаж и финансового планирования.

# Surgery Revenue Forecast

> Python application for forecasting future surgery revenue based on treatment plans.

## Description

This project analyzes scheduled surgeries together with treatment plans stored in the corporate CRM database.

Since each patient may have several treatment plans before surgery, the application calculates both the minimum and maximum expected revenue for every planned operation. This allows management to estimate a realistic revenue range instead of relying on a single value.

Important:

The project requires access to the hospital database.

Without the original database structure the application cannot be executed.

## Business Goal

The application helps forecast future surgery revenue using treatment plans linked to scheduled operations.

The generated report was used for financial planning and revenue forecasting.

## Features

- SQL database integration
- Surgery schedule analysis
- Treatment plan analysis
- Revenue forecasting
- Minimum revenue calculation
- Maximum revenue calculation
- Duplicate operation detection
- Excel report generation

## Tech Stack

- Python
- pandas
- NumPy
- SQLAlchemy
- Firebird SQL
- xlsxwriter
- python-dotenv
- pyTelegramBotAPI

## How It Works

1. Loads surgery schedule
2. Loads treatment plans
3. Matches patients with treatment plans
4. Calculates minimum planned revenue
5. Calculates maximum planned revenue
6. Removes duplicate operations
7. Generates Excel report

## Example / Demo

### Input

Hospital database

- Surgery schedule
- Treatment plans
- Patient information

### Output

Excel report containing:

- Planned surgeries
- Minimum expected revenue
- Maximum expected revenue
- Historical surgeries

## Use Case

This project can be used for:

- revenue forecasting
- healthcare analytics
- financial planning
- business analytics
- reporting
