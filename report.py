# Define columns based on user instructions
sector_column = 'What is the Sector?'
mda_column = 'What is the name of the MDA in charge?'
project_title_column = 'What project is this response for?'

# Columns for detailed information
general_info_columns = ['Title of Project:', 'Description of Total Achievement Recorded Since Inception', 'Project/ Programme Cost Variation if any?', 
                        'What is the total project cost?', 'Project/ Programme Cost Variation if any?', 'Percentage of Total Achievement Recorded Since Inception',
                        'Geo-Political Zone', 'What is the project start date?', 'What is the project end date?']

location_columns = ['State Located ?', 'Where is this project located?', 'Upload project images']
financial_performance_columns = ['What is the total project cost?', 'What is the total amount appropriated this year?', 
                                 'Percentage of Total Achievement Recorded Since Inception']
result_delivery_columns = ['Deliverables', 'Key Performance Indicators']
lesson_learned_columns = ['What are the project challenges?', 'What is the project status?']
recommendations_column = 'What are your recommendations?'

# Prepare the data by categorizing by Sector and MDA
report_data = []

# Group by 'Sector' and then by 'MDA'
for sector, sector_group in excel_data.groupby(sector_column):
    sector_reports = []
    for mda, mda_group in sector_group.groupby(mda_column):
        mda_reports = []
        for _, row in mda_group.iterrows():
            # General Information Section
            general_info = f"Title of Project: {row[project_title_column]}\n"
            general_info += f"{row['Description of Total Achievement Recorded Since Inception']}.\n"
            general_info += f"Start Date: {row['What is the project start date?']} - End Date: {row['What is the project end date?']}.\n"

            # Location Information
            location_info = f"State: {row['State Located ?']}, Location Coordinates: {row['Where is this project located?']}\n"

            # Situational Analysis - Financial Performance
            financial_performance = f"Project Cost: {row['What is the total project cost?']}, Appropriated Amount This Year: {row['What is the total amount appropriated this year?']}, Percentage Completion: {row['Percentage of Total Achievement Recorded Since Inception']}%.\n"

            # Situational Analysis - Results Delivery
            results_delivery = f"Deliverables: {row['Deliverables']}, Key Performance Indicators: {row['Key Performance Indicators']}.\n"

            # Lessons Learned
            lessons_learned = f"Challenges: {row['What are the project challenges?']}, Status: {row['What is the project status?']}.\n"

            # Recommendations
            recommendations = f"Recommendations: {row[recommendations_column]}.\n"

            # Compile the report for each project under the MDA
            mda_report = f"General Information:\n{general_info}\nLocation Information:\n{location_info}\n"
            mda_report += f"Situation Analysis:\nFinancial Performance:\n{financial_performance}\nResults Delivery:\n{results_delivery}\n"
            mda_report += f"Lessons Learned:\n{lessons_learned}\nRecommendations:\n{recommendations}\n"

            # Append each MDA report
            mda_reports.append(mda_report)

        # Append the reports for each MDA
        sector_reports.append(f"Sector: {sector}\nMDA: {mda}\n" + "\n".join(mda_reports) + "\n")

    # Append the complete sector report
    report_data.append("\n".join(sector_reports))

# Combine all the sector reports into one comprehensive report
comprehensive_report = "\n".join(report_data)

# Display the first part of the comprehensive report to check
comprehensive_report[:2000]  # Displaying a snippet due to length limitations
