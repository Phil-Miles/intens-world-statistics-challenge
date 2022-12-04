# I am going to use the pandas module to open the .xlsx file
import pandas
df = pandas.read_excel("WSC Input.xlsx")

# this is the dictionary that I will convert to pandas data frame and then export as .csv
output = {"Id": [], "Word": [], "Total Count": [], "Single Chapter Max-Count": [], "Unique Chapter Count": [],
                    "Single Heading Max-Count": [], "Unique Heading Count": [], "Single Subheading Max-Count": [],
                    "Unique Subheading Count": [], "Single Duty Rate Max-Count": [], "Unique Duty Rate Count": [],
                    "Single Tariff Max-Count": [], "Unique Tariff Count": []}
# parsed_words is going to hold every single word from the entire "Description" column
parsed_words = []
# tariff_list is going to hold every single Tariff Number for each word I am parsing
tariff_list = []
# unique_words will consist of one instance of each occurring word in the "Description" column
unique_words = []

# I am looping through the rows of the imported dataframe
for entry in df.index:
    content = df["Description"][entry]
    tariff_number = df["TariffNumber"][entry]
    tariff = str(tariff_number)
    description = content.lower().split()
    # each word is being appended to the parsed_words list
    for word in description:
        parsed_words.append(word)
        # each tariff number (with "0" in front if needed) is being appended to the tariff_list
        if len(tariff) < 10:
            tariff = "0" + tariff
        tariff_list.append(tariff)
        if word not in unique_words:
            # unique_words list purpose is twofold:
            # 1. to form a "Word" column in the output .csv
            # 2. to allow us to loop through every used word once, no matter how many occurrences it has
            unique_words.append(word)

# keyword_tariffs will hold every tariff where a certain word has occurred
keyword_tariffs = []
# unique_keyword_tariffs will hold one instance of each tariff the word has occurred at
unique_keyword_tariffs = []
# word_id is going to increase by 1 for each word we enter into the output .csv file
word_id = 1000


# max_count function returns the Max-Count value for each sub-tariff (and tariff) for each word
# t_length determines how short/long the sub-tariff is: 2 for chapter, 4 for heading, 6 for sub-heading, etc.
def max_count(t_length):
    global keyword_tariffs
    sub_tariff_list = [tar[:t_length] for tar in keyword_tariffs]
    occurrences = []
    for tar in range(len(sub_tariff_list)):
        counter = sub_tariff_list.count(sub_tariff_list[tar])
        occurrences.append(counter)
    return max(occurrences)


# unique_count counts unique occurrences of each tariff where the word has occurred
# there are two functions for easier readability and explanation
# one function with "mode" attribute would to the job of both
def unique_count(t_length):
    global unique_keyword_tariffs
    sub_tariff_list = [tar[:t_length] for tar in keyword_tariffs]
    unique_keyword_tariffs = []
    for occurrence in range(len(sub_tariff_list)):
        if sub_tariff_list[occurrence] not in unique_keyword_tariffs:
            unique_keyword_tariffs.append(sub_tariff_list[occurrence])
    return len(unique_keyword_tariffs)


# now that I have a proper set-up, I will loop through the unique_words list
for word in unique_words:
    for parsed_word in range(len(parsed_words)):
        if word == parsed_words[parsed_word]:
            # I will append all the tariffs where the word occurs to keyword_tariffs
            keyword_tariffs.append(tariff_list[parsed_word])
            for tariff in keyword_tariffs:
                if tariff not in unique_keyword_tariffs:
                    # I will append all the unique tariffs for this keyword to the unique_keyword_tariffs
                    unique_keyword_tariffs.append(tariff)
    # finally, I will prepare all the data we need for the entire row of the dataframe
    data = [word_id, word, len(unique_keyword_tariffs), max_count(2), unique_count(2), max_count(4), unique_count(4),
            max_count(6), unique_count(6), max_count(8), unique_count(8), max_count(10), unique_count(10)]
    data_index = 0
    # I will loop through the output dictionary and append all the values, in their correct and expected order
    for key in output:
        output[key] += [data[data_index]]
        data_index += 1
    # this is the moment where the two "placeholder" lists are being wiped and the word_id is being increased by 1
    keyword_tariffs = []
    unique_keyword_tariffs = []
    word_id += 1
# after all the words have been looped through, we can create a pandas data frame out of the output dictionary
export = pandas.DataFrame(output)
# and export the dataframe as a .csv file
export.to_csv("output.csv")
