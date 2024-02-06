import spacy
import pandas as pd
import time

# Load spaCy model
nlp = spacy.load("en_core_web_trf")

# Read Excel file using pandas
excel_file_path = "/Users/aiman/Documents/People/EishiBishi/parts_of_speech_tagger/ThesisData4.xlsx"

try:
    df = pd.read_excel(excel_file_path)
except Exception as e:
    error_message = f"Error reading Excel file: {e}"
    print(error_message)

    # Save the error information to a log file
    log_path = "/Users/aiman/Documents/People/EishiBishi/parts_of_speech_tagger/error_log.txt"
    with open(log_path, 'a') as log_file:
        log_file.write(error_message + '\n')

    # Create an empty DataFrame to avoid issues later in the script
    df = pd.DataFrame()

# Define a function to process text with spaCy
def process_text(text):
    # Check if text is not NaN
    if pd.notna(text):
        doc = nlp(text)
        results = []
        for token in doc:
            results.append([token.text, token.lemma_, token.tag_])
        return results
    else:
        return []

# Create a DataFrame to store POS counts for each post
columns = [
    'coordinating_conjunction', 'cardinal_number', 'determiner',
    'existential_there', 'foreign_word', 'preposition_sub_conj',
    'that_as_subordinator', 'adjective', 'adjective_comparative',
    'adjective_superlative', 'list_marker', 'modal', 'noun_singular_mass',
    'noun_plural', 'proper_noun_singular', 'proper_noun_plural',
    'predeterminer', 'possessive_ending', 'personal_pronoun',
    'possessive_pronoun', 'adverb', 'adverb_comparative',
    'adverb_superlative', 'particle', 'sentence_break_punctuation',
    'symbol', 'infinitive_to', 'interjection', 'verb_base_form',
    'verb_past_tense', 'verb_gerund_present_participle',
    'verb_past_participle', 'verb_sing_present_non_3d',
    'verb_3rd_person_sing_present', 'wh_determiner', 'wh_pronoun',
    'possessive_wh_pronoun', 'wh_adverb', 'hash', 'dollar',
    'quotation_marks', 'opening_quotation_marks', 'opening_brackets',
    'closing_brackets', 'comma', 'punctuation'
]

pos_counts_df = pd.DataFrame(columns=columns)

# Process each row in the DataFrame
start_time = time.time()  # Start the timer
try:
    for index, row in df.iterrows():
        text = row["contents"]

        # Process text with spaCy
        processed_text = process_text(text)

        # Count occurrences of each POS tag for the current post
        post_tag_counts = pd.DataFrame(processed_text, columns=['Text', 'Lemma', 'Tag'])['Tag'].value_counts().reset_index()
        post_tag_counts.columns = ['POS_Tag', 'Count']

        # Transpose the DataFrame to have counts horizontally aligned
        post_counts_row = post_tag_counts.set_index('POS_Tag').T
        post_counts_row['index'] = index

        # Append the counts to the pos_counts_df
        pos_counts_df = pd.concat([pos_counts_df, post_counts_row], ignore_index=True)

        # Print a progress statement after every 500 rows
        if index % 1000 == 0 and index > 0:
            elapsed_time = time.time() - start_time
            print(f"Processed {index} rows. Elapsed time: {elapsed_time:.2f} seconds.")
except Exception as e:
    print(f"Error during processing: {e}")

# Merge the original DataFrame with the pos_counts_df
df_with_counts = pd.merge(df, pos_counts_df, left_index=True, right_on='index', how='left')

# Drop unnecessary columns (index) if they exist
df_with_counts = df_with_counts.drop(['index'], axis=1, errors='ignore')

# Save the merged DataFrame to a new Excel file
output_excel_path = "/Users/aiman/Documents/People/EishiBishi/parts_of_speech_tagger/parts_of_speech_results4.xlsx"
try:
    df_with_counts.to_excel(output_excel_path, index=False)
    elapsed_time = time.time() - start_time
    print(f"Processed data with post-level tag counts saved to {output_excel_path}. Total elapsed time: {elapsed_time:.2f} seconds.")
except Exception as e:
    print(f"Error saving output Excel file: {e}")
