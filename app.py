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

        # Interpret the data
        interpret_event_data(event_details, event_type)
        interpret_income_data(income_details, projections)
    else:
        st.warning('Stock symbol not found in the data. Please check the symbol and try again.')

# Function to interpret event data
def interpret_event_data(details, event_type):
    st.write("### Interpretation of Event Data")
    if 'Event Coefficient' in details.index:
        if details['Event Coefficient'] < -1:
            st.write(f"**1% Increase in {event_type}:** Stock price decreases significantly. Increase portfolio risk.")
        elif details['Event Coefficient'] > 1:
            st.write(f"**1% Increase in {event_type}:** Stock price increases, benefiting from {event_type.lower()}.")

# Function to interpret income data
def interpret_income_data(details, projections):
    st.write("### Interpretation of Income Statement Data")
    
    correlation_ranges = {
        'Very Good': (0.80, 1.00),
        'Good': (0.60, 0.79),
        'Neutral': (0.40, 0.59),
        'Bad': (0.20, 0.39),
        'Very Bad': (0.00, 0.19),
    }

    for col in projections['Parameter']:
        if col in details.index:
            current_value = pd.to_numeric(details[col], errors='coerce')
            projected_value = pd.to_numeric(projections[projections['Parameter'] == col]['Projected Value'], errors='coerce').values[0]
            
            # Sample interpretation based on correlation (replace with actual correlation calculation if available)
            correlation = (projected_value - current_value) / current_value  # Placeholder for correlation
            
            for key, (lower, upper) in correlation_ranges.items():
                if lower <= correlation < upper:
                    st.write(f"**{col}:** {key} correlation with projected changes.")
                    st.write(get_detailed_interpretation(col, key))
                    break

# Function to get detailed interpretation based on correlation
def get_detailed_interpretation(metric, range_label):
    interpretations = {
        'Total Revenue/Income': {
            'Very Good': "Government initiatives spur growth.",
            'Good': "Increased consumer spending drives revenue.",
            'Neutral': "Market saturation limits revenue growth.",
            'Bad': "Economic downturn affects sales.",
            'Very Bad': "Structural issues lead to revenue decline."
        },
        'Total Operating Expense': {
            'Very Good': "Efficient cost management aligns with growth.",
            'Good': "Inflation impacts costs but is manageable.",
            'Neutral': "Cost controls vary across sectors.",
            'Bad': "Rising raw material costs squeeze margins.",
            'Very Bad': "Inefficiency leads to ballooning costs."
        },
        'Operating Income/Profit': {
            'Very Good': "High correlation indicates robust margins.",
            'Good': "Growth strategies enhance profitability.",
            'Neutral': "Margins affected by competition.",
            'Bad': "High fixed costs reduce operating income.",
            'Very Bad': "Non-competitive sectors suffer losses."
        },
        'EBITDA': {
            'Very Good': "Strong operational efficiency.",
            'Good': "Positive cash flow trends support growth.",
            'Neutral': "Limited growth due to external factors.",
            'Bad': "Declining sectors struggle with EBITDA.",
            'Very Bad': "Poor management leads to negative EBITDA."
        },
        'EBIT': {
            'Very Good': "Reflects operational excellence.",
            'Good': "Investment in innovation boosts EBIT.",
            'Neutral': "Mixed results due to competition.",
            'Bad': "Regulatory pressures impact EBIT margins.",
            'Very Bad': "Companies unable to adapt face EBIT losses."
        },
        'Income/Profit Before Tax': {
            'Very Good': "Tax incentives and growth correlate.",
            'Good': "Profitability rises with economic conditions.",
            'Neutral': "Varying tax policies create unpredictability.",
            'Bad': "Unfavorable tax conditions reduce profits.",
            'Very Bad': "Structural inefficiencies lead to losses."
        },
        'Net Income From Continuing Operation': {
            'Very Good': "Sustained growth reflects stability.",
            'Good': "Positive shifts in economic policies boost net income.",
            'Neutral': "Fluctuations in income stability.",
            'Bad': "High debt levels hamper net income.",
            'Very Bad': "Companies may face losses in core operations."
        },
        'Net Income': {
            'Very Good': "Economic reforms boost profitability.",
            'Good': "Strong demand fuels net income growth.",
            'Neutral': "External factors affect overall profitability.",
            'Bad': "High competition diminishes net income.",
            'Very Bad': "Losses indicate critical issues in operations."
        },
        'Net Income Applicable to Common Share': {
            'Very Good': "High returns drive investor confidence.",
            'Good': "Shareholder value grows with profitability.",
            'Neutral': "Fluctuations in earnings affect share prices.",
            'Bad': "Weak performance leads to dividend cuts.",
            'Very Bad': "Poor performance reflects systemic issues."
        },
        'EPS (Earning Per Share)': {
            'Very Good': "Strong EPS growth attracts investors.",
            'Good': "Moderate growth signals healthy business.",
            'Neutral': "EPS stability amidst market fluctuations.",
            'Bad': "Low EPS growth raises concerns.",
            'Very Bad': "Negative EPS reflects deep-rooted issues."
        }
    }
    return interpretations.get(metric, {}).get(range_label, "No interpretation available.")

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
            explanation = "Dynamic calculation considers the event coefficient and rate change."
        else:  # Simple
            price_change = latest_close_price * (expected_rate / 100)
            projected_price = latest_close_price + price_change
            change = expected_rate
            explanation = "Simple calculation uses the expected rate directly."

        new_row = pd.DataFrame([{
            'Parameter': 'Projected Stock Price',
            'Current Value': latest_close_price,
            'Projected Value': projected_price,
            'Change': change
        }])
        projections = pd.concat([projections, new_row], ignore_index=True)
    else:
        st.warning("Stock Price data not available in event details.")

    # Project changes in new income statement items
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
                
                projections = projections.append({
                    'Parameter': column,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': projected_value - current_value
                }, ignore_index=True)

    return projections

# Run the app
if __name__ == '__main__':
    get_stock_details(stock_name, event_type, calculation_method)
