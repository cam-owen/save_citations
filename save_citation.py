# Import the necessary modules.
import webbrowser  # module for web searching
import re  # module for having multiple delimiters in the split method
import openpyxl  # module for editing excel files

# Make sure you:
# 1. Create an excel sheet ahead.
# 2. Input "Author(s)", "Year", and "Title" in the 1st 3 columns of the 1st row.
# 3. Update the name of your excel file here before running the program.
excel_file = 'references.xlsx'
# 4. Keep the excel file closed. The program does not update the excel file when it is open.


class GetArticleApp:
    def __init__(self):
        # Step 1: Insert citation into the terminal. Works best if you include only author, year, and title
        # and exclude journal outlet or book publisher.
        ref = input(f"To store, please insert citation here: ")
        self.reference = ref.translate({ord('.'): None})
        # Step 2: Split the provided citation into a list of three components - author(s), year, and title.
        # NOTE: The app splits the citation through the "year", so if there's a number in the title,
        # things can be a bit wonky. But it does not affect the other entries.
        self.li = re.split('([0-9]+)', self.reference)
        print(self.li)

    # Step 3: Store author(s), year, and title into Excel sheet.
    def insert_data_into_row(self):
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
        sheet.append(self.li)
        workbook.save(filename=excel_file)

    # Step 4: Search for citation in Google Scholar.
    def search_article_info(self):
        prompt = input(f"Would you like to search for the citation on Google Scholar? "
                       f"Please answer yes or no: ").lower()
        if prompt == "yes":
            webbrowser.open(url=f'https://scholar.google.com/scholar?hl=en&as_sdt=0%2C5&q={self.reference}&btnG=&oq=',
                            new=0, autoraise=True)
        else:
            print("Sounds good.")


# Step 5: Decide whether we want to input a new reference.
def play_again():
    return input("Would you like to store another reference? Please answer yes or no: ").lower() == "yes"


# If yes, the program starts over.
# If no, the program ends.
while True:
    article = GetArticleApp()
    article.insert_data_into_row()
    article.search_article_info()
    if not play_again():
        break


print("Program ended.")
