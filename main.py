from checker import DataChecker
from writer import DataResultWriter

checker = DataChecker()
checker.check_documents()
writer = DataResultWriter(checker.result_df)
writer.write_result()

