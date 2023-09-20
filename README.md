# VBA Challenge

## Assignment Overview

In this VBA challenge activity, we aim to create a new automated process using VBA (Visual Basic for Applications) to loop through a dataset containing stock information for a year and extract specific information for each stock. The key tasks include:

### Step 1: Identify Ticker Symbols

As we iterate through the rows of data, we will look for changes in the ticker symbol in the first column. These changes will serve as our markers to identify the range for each specific stock (ticker).

### Step 2: Calculate Key Metrics

Using logical operations, we will extract values from the most recent change in ticker to the next change in ticker. This will allow us to calculate the opening and closing values for each stock by taking the difference between the two.

### Step 3: Determine Yearly Change and Percentage Change

To assess stock performance, we will employ a color-coding system. We will designate whether the yearly change is positive (green) or negative (red). Additionally, we will calculate the percentage change for each stock by dividing the change in value by the starting value for that particular ticker.

### Step 4: Calculate Total Stock Volume

While looping through the ticker symbols, we will continuously add up the trading volume. We will reset this volume count to zero each time there is a change in ticker, effectively accumulating the total stock volume over the year.

### Step 5: Identify Maximum Values

Throughout the process, we will track and update three key metrics: MaxUp (maximum positive percentage change), MaxDown (maximum negative percentage change), and MaxVol (maximum total stock volume). These values will start at zero and be replaced as we iterate through the dataset, identifying larger or lower values.

---

## Detailed Steps

1. **Read the File**: Begin by reading the dataset file.

2. **Create New Model**: Implement a new VBA model to automate the steps described above. This model will involve looping through the dataset, identifying ticker symbols, calculating key metrics, color-coding results, and tracking maximum values.

---

This VBA challenge activity involves creating an efficient and automated process to analyze stock data for a given year. By following the steps outlined above, we can extract valuable insights from the dataset, including yearly changes, percentage changes, and total trading volumes, while also identifying top-performing and underperforming stocks.

