# **Customer_Segmentation_RFM_Analysis**
## Objective
Management requested a method to rank and group customers based on their purchasing behaviour. The goal is to segment customers based on their value and engagement levels to identify high-value customers, at-risk customers, and those needing re-engagement.

Customer Segmentation Technique : 
Customers were ranked and grouped using the RFM model (Recency, Frequency, Monetary), which evaluates the recency of their last purchase, how frequently they purchase, and their total monetary value. This allows for identifying the best customers and targeting them in marketing campaigns and also helps identifying at risk customers who require immediate attention.

### Prerequisites
- MS-Excel
- MS-PowerPoint
<br />

## Data and Methodology
### 1. Data Preparation
 - #### Pivot Table Creation:
    - A Pivot Table was created based on the Customer ID to aggregate the data:
          
 - #### Frequency Calculation:
    - Frequency was determined by counting the number of purchases made by each customer, using the `InvoiceNo` as an indicator.
    
 - #### Recency Calculation:
    - Recency was derived by finding the maximum invoice date for each customer, representing the most recent purchase.
    - The DATEDIF() function, which is an archived Excel function, was used to calculate the difference between the current date and the most recent purchase date.
    - Syntax for DATEDIF():
       ``` DATEDIF(start_date, end_date, unit) ```
    
 - #### Monetary Calculation:
    - The monetary value was calculated by summing the total amount spent by each customer.
    
### 2. RFM Scoring
- #### Frequency & Monetary Scores:
    - The PERCENTRANK.INC() function was applied to calculate the Frequency and Monetary scores.
    - This function calculates the relative rank of a value within a data set as a percentage between 0 and 1.
    - In order to make the score out of 10, we multiply the output by 10 to get the final score.
    

- #### Recency Score:
    - Unlike other scores, to assign a higher score to more recent purchases, the formula was flipped as:
    
    ```=(1 - PERCENTRANK.INC(array, value, significance))```  Or 
    ``` =(1 - PERCENTRANK.INC($F$2:$F$4375,F2,1))```
      
- #### RFM Score:
    - The overall RFM total was computed by summing the Recency, Frequency, and Monetary scores.
    - Again, The PERCENTRANK.INC() function was applied to the RFM Total Column and the output was multiplied by 10 to get the Final RFM Score which we will use for customer segmentation.

### 3. Customer Segmentation
- #### Segmentation Criteria:
    - Customers were segmented into categories based on their RFM scores:
        - ##### Top Customers (Score: 10-8) : High scores across Recency, Frequency, and Monetary metrics.
        - ##### Loyal Customers (Score: 7-5) : Consistent purchasers with moderate spending.
        - ##### At-Risk Customers (Score: 4-2) : Customers with low recency and spending, requiring re-engagement.
        - ##### Immediate Attention Customers (Score: 1-0) : Customers with recent activity but low frequency and spending, needing targeted follow-up to boost engagement.
          
- #### Customer Segments Filtering:
    - VLOOKUP() function was used to assign customers to specific segments based on their RFM score.
    - Additionally, By applying filters on the RFM score column, we can classify the Top Customers & At Risk / Immediate Attention Customers by selecting top values (e.g. 10, 9, 8) and lower values (e.g.4,3,2,1,0)
      

### 4. Handling Missing Data:
- Customers with missing IDs and associated data were excluded to prevent skewed analysis results.
<br />
      
## Key Insights
- #### **Top Customers :** Represent 18% of the customer base, with an average spending of $7,062 and an average last purchase within 15 days.
- #### **At Risk & Immediate Attention Customers :** Comprise 51% of the customer base, spending an average of $379, with an average last purchase 162 days ago.

  ![]
## Recommendations
- #### **Targeted Campaigns :**
    -   Develop re-engagement strategies with tailored incentives for at-risk customers.
- #### **Top Customer Retention :**
    -  Analyze top and loyal customers further to explore opportunities for increasing their lifetime value.
