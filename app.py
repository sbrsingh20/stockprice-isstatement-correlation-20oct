import pandas as pd
import streamlit as st

# Load the data from Excel files
inflation_data = pd.read_excel('Inflation_event_stock_analysis_resultsOct.xlsx')
income_data = pd.read_excel('Inflation_IncomeStatement_correlation_results.xlsx')
interest_rate_data = pd.read_excel('interestrate_event_stock_analysis_resultsOct.xlsx')
interest_rate_income_data = pd.read_excel('interestrate_IncomeStatement_correlation_results.xlsx')

# Set up Streamlit app
st.title('Stock Analysis Based on Economic Events')

# Create a sidebar for user input
st.sidebar.header('Search for a Stock')
event_type = st.sidebar.selectbox('Select Event Type:', ['Inflation', 'Interest Rate'])
stock_name = st.sidebar.text_input('Enter Stock Symbol:', '')
expected_rate = st.sidebar.number_input('Enter Expected Upcoming Rate (%):', value=3.65, step=0.01)
calculation_method = st.sidebar.selectbox('Select Calculation Method:', ['Dynamic', 'Simple'])

# Function to fetch details for a specific stock based on the event type
def get_stock_details(stock_symbol, event_type, method):
    if event_type == 'Inflation':
        event_row = inflation_data[inflation_data['Symbol'] == stock_symbol]
        income_row = income_data[income_data['Stock Name'] == stock_symbol]
    else:  # Interest Rate
        event_row = interest_rate_data[interest_rate_data['Symbol'] == stock_symbol]
        income_row = interest_rate_income_data[interest_rate_income_data['Stock Name'] == stock_symbol]

    if not event_row.empty and not income_row.empty:
        event_details = event_row.iloc[0]
        income_details = income_row.iloc[0]

        st.subheader(f'Details for {stock_symbol}')
        
        # Display event data
        st.write(f"### {event_type} Event Data")
        st.write(event_row)

        # Display income statement data
        st.write("### Income Statement Data")
        st.write(income_row)

        # Generate projections based on expected rate
        projections = generate_projections(event_details, income_details, expected_rate, event_type, method)
        
        # Display projections
        st.write(f"### Projected Changes Based on Expected {event_type}")
        st.dataframe(projections)

        # Interpretations
        interpret_data(event_details, income_details, event_type)
    else:
        st.warning('Stock symbol not found in the data. Please check the symbol and try again.')

# Function to interpret data
def interpret_data(event_details, income_details, event_type):
    if event_type == 'Inflation':
        interpret_inflation_data(event_details)
    else:
        interpret_interest_rate_data(event_details)

    interpret_income_data(income_details)

# Function to interpret inflation data
def interpret_inflation_data(details):
    st.write("### Interpretation of Inflation Event Data")
    if 'Event Coefficient' in details.index:
        if details['Event Coefficient'] < -1:
            st.write("**1% Increase in Inflation:** Stock price decreases significantly. Increase portfolio risk.")
        elif details['Event Coefficient'] > 1:
            st.write("**1% Increase in Inflation:** Stock price increases, benefiting from inflation.")

# Function to interpret interest rate data
def interpret_interest_rate_data(details):
    st.write("### Interpretation of Interest Rate Event Data")
    if 'Event Coefficient' in details.index:
        if details['Event Coefficient'] < -1:
            st.write("**1% Increase in Interest Rate:** Stock price decreases significantly. Increase portfolio risk.")
        elif details['Event Coefficient'] > 1:
            st.write("**1% Increase in Interest Rate:** Stock price increases, benefiting from interest hikes.")

# Function to interpret income data
def interpret_income_data(details):
    st.write("### Interpretation of Income Statement Data")
    
    correlations = {
        'Total Revenue/Income': details.get('Total Revenue/Income Correlation', 0),
        'Total Operating Expense': details.get('Total Operating Expense Correlation', 0),
        'Operating Income/Profit': details.get('Operating Income/Profit Correlation', 0),
        'EBITDA': details.get('EBITDA Correlation', 0),
        'EBIT': details.get('EBIT Correlation', 0),
        'Income/Profit Before Tax': details.get('Income/Profit Before Tax Correlation', 0),
        'Net Income From Continuing Operation': details.get('Net Income From Continuing Operation Correlation', 0),
        'Net Income': details.get('Net Income Correlation', 0),
        'Net Income Applicable to Common Share': details.get('Net Income Applicable to Common Share Correlation', 0),
        'EPS (Earning Per Share)': details.get('EPS Correlation', 0),
    }

    for metric, correlation in correlations.items():
        interpretation = interpret_correlation(correlation)
        st.write(f"**{metric}:** {interpretation}")

def interpret_correlation(correlation):
    if correlation > 0.8:
        return "Very Good: Indicates a strong positive relationship. Economic events significantly influence these income items."
    elif correlation > 0.6:
        return "Good: Reflects a moderate to strong correlation. Economic growth impacts revenue positively."
    elif correlation > 0.4:
        return "Neutral: Shows a moderate correlation. Other factors may play a larger role."
    elif correlation > 0.2:
        return "Bad: Indicates a weak correlation. Companies might need to re-evaluate their strategies."
    else:
        return "Very Bad: Shows little to no correlation. Economic events do not significantly affect income items."

# Function to generate projections based on expected rate and calculation method
def generate_projections(event_details, income_details, expected_rate, event_type, method):
    latest_event_value = pd.to_numeric(income_details.get('Latest Event Value', 0), errors='coerce')
    projections = pd.DataFrame(columns=['Parameter', 'Current Value', 'Projected Value', 'Change'])

    if 'Latest Close Price' in event_details.index:
        latest_close_price = pd.to_numeric(event_details['Latest Close Price'], errors='coerce')

        if method == 'Dynamic':
            rate_change = expected_rate - latest_event_value
            price_change = event_details['Event Coefficient'] * rate_change
            projected_price = latest_close_price + price_change
            change = price_change
        else:  # Simple
            price_change = latest_close_price * (expected_rate / 100)
            projected_price = latest_close_price + price_change
            change = expected_rate

        new_row = pd.DataFrame([{
            'Parameter': 'Projected Stock Price',
            'Current Value': latest_close_price,
            'Projected Value': projected_price,
            'Change': change
        }])
        projections = pd.concat([projections, new_row], ignore_index=True)

    # Project changes in income statement items
    for column in income_details.index:
        if column != 'Stock Name':
            current_value = pd.to_numeric(income_details[column], errors='coerce')
            if pd.notna(current_value):
                if method == 'Dynamic':
                    if 'Income Coefficient' in event_details.index:
                        change = current_value * (expected_rate / 100) * event_details['Income Coefficient']
                        projected_value = current_value + change
                    else:
                        projected_value = current_value
                else:  # Simple
                    projected_value = current_value + (current_value * (expected_rate / 100))
                
                new_row = pd.DataFrame([{
                    'Parameter': column,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': projected_value - current_value
                }])
                projections = pd.concat([projections, new_row], ignore_index=True)

    return projections

# Check if user has entered a stock symbol and selected an event
if stock_name and event_type:
    get_stock_details(stock_name, event_type, calculation_method)

