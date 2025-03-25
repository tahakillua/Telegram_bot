import pandas as pd
import seaborn as sns
import arabic_reshaper
import matplotlib.pyplot as plt
from bidi.algorithm import get_display

# Column
ID = "رقم التعريف"
FAMILY_NAME = "اللقب"
FIRST_NAME = "الاسم"
BIRTH_DAY = "تاريخ الميلاد"
EVALUATION = "معدل تقويم النشاطات /20"
ASSIGNMENT = "الفرض /20"
EXAM = "الإختبار /20"
SCORE = "معدل المادة"
SCORE_RANGE = "الكفاءة"
DISTRIBUTION_OF_STUDENTS_SCORE = "توزيع درجات الطلاب"
TOTAL_STUDENTS = "عدد التلاميذ"
BOX_PLOT_OF_STUDENT_SCORE = "مخطط الصندوق لدرجات الطلاب"
AVERAGE_SCORES_EV_ASS_EX = "متوسط الدرجات عبر التقييمات المختلفة"
ASSESSMENT_TYPE = "نوع التقييم"
AVERAGE_SCORE = "متوسط الدرجات"
EVALUATION_VS_EXAM_SCORE = "تقييم الدرجات مقابل درجات الاختبار"
CATEGORY_RATION = "نسبة الفئة"
PERCENTAGE_DISTRIBUTION_OF_SCORE = "التوزيع النسبي للدرجات"
CORRELATION_BETWEEN_ASSESSMENT = "الارتباط بين التقييمات"

# Normalisation the arabic sentence for graph
SCORELBL = get_display(arabic_reshaper.reshape(SCORE))
TOTALSUTENDTLBL = get_display(arabic_reshaper.reshape(TOTAL_STUDENTS))
TITLELBL = get_display(arabic_reshaper.reshape(DISTRIBUTION_OF_STUDENTS_SCORE))
BOXPLOTOFSTUDENTSCORELBL = get_display(arabic_reshaper.reshape(BOX_PLOT_OF_STUDENT_SCORE))
AVERAGESCORESEVASSEXLBL = get_display(arabic_reshaper.reshape(AVERAGE_SCORES_EV_ASS_EX))
ASSESSMENTTYPELBL = get_display(arabic_reshaper.reshape(ASSESSMENT_TYPE))
AVERAGESCORELBL = get_display(arabic_reshaper.reshape(AVERAGE_SCORE))
EVALUATIONVSEXAMSCORELBL = get_display(arabic_reshaper.reshape(EVALUATION_VS_EXAM_SCORE))
EVALUATIONLBL = get_display(arabic_reshaper.reshape(EVALUATION))
EXAMLBL = get_display(arabic_reshaper.reshape(EXAM))
CATEGORYRATIONLBL = get_display(arabic_reshaper.reshape(CATEGORY_RATION))
PERCENTAGEDISTRIBUTIONOFSCORELBL = get_display(arabic_reshaper.reshape(PERCENTAGE_DISTRIBUTION_OF_SCORE))
CORRELATIONBETWEENASSESSMENTLBL = get_display(arabic_reshaper.reshape(CORRELATION_BETWEEN_ASSESSMENT))

# Category Score Of Students
bins = [0, 5, 10, 15, 20]
labels = ["د", "ج", "ب", "أ"]


# Calcul The Score Of The Subject
def rate(x, y, z):
    return (x * 2 + (y + z) / 2) / 3


class AnalyseVisualizationFile:
    def __init__(self, path, sheet):
        self.file = pd.read_excel(path, sheet_name=sheet, skiprows=7)
        self.file[SCORE] = pd.Series([rate(row[EXAM], row[ASSIGNMENT], row[EVALUATION])
                                      for (index, row) in self.file.iterrows()])
        self.analysedata = {}
        self.analysetable = None

    def numberstudent(self):
        total_students = self.file.shape[0]
        self.analysedata["[total_students]"] = total_students

    def overallaverage(self):
        overall_average = self.file[SCORE].mean()
        self.analysedata["[overall_average]"] = round(overall_average, 2)

    def percentagedistribution(self):
        self.file[SCORE_RANGE] = pd.cut(self.file[SCORE], bins=bins, labels=labels, right=False)
        percentage_distribution = self.file[SCORE_RANGE].value_counts(normalize=True, sort=False) * 100
        self.analysetable = percentage_distribution

    def standarddeviation(self):
        standard_deviation = self.file[SCORE].std()
        self.analysedata["[standard_deviation]"] = round(standard_deviation, 2)

    def studenthighscore(self):
        index_of_max_score = self.file[SCORE].idxmax()
        student_high_score_inf = self.file.loc[index_of_max_score]
        student_high_score_name = student_high_score_inf[FAMILY_NAME] + " " + student_high_score_inf[FIRST_NAME]
        high_score = student_high_score_inf[SCORE]
        self.analysedata["[student_high_score]"] = student_high_score_name
        self.analysedata["[high_score]"] = round(high_score, 2)

    def percentage(self):
        total_students = self.file.shape[0]
        students_with_10_or_more = self.file[self.file[SCORE] >= 10].shape[0]
        percentage = (students_with_10_or_more / total_students)*100
        self.analysedata["[percentage_pass]"] = round(percentage, 2)

    def histogram(self, path):
        # Plotting the histogram of scores
        plt.hist(self.file[SCORE], bins=10, edgecolor='black')
        plt.title(TITLELBL)
        plt.xlabel(SCORELBL, fontdict=None, labelpad=None)
        plt.ylabel(TOTALSUTENDTLBL, fontdict=None, labelpad=None)
        plt.grid(True)
        plt.savefig(path)
        plt.close()

    def boxplot(self, path):
        # Plotting the box plot of scores
        plt.boxplot(self.file[SCORE], vert=False)
        plt.title(BOXPLOTOFSTUDENTSCORELBL)
        plt.xlabel(SCORELBL)
        plt.grid(True)
        plt.savefig(path)
        plt.close()


    def barchar(self, path):
        # Plotting the bar chart for average scores
        # Calculate average scores for each type
        average_scores = self.file[[EVALUATION, ASSIGNMENT, EXAM, SCORE]].mean()
        average_scores.plot(kind='bar', color=['skyblue', 'orange', 'green', 'red'], )
        plt.title(AVERAGESCORESEVASSEXLBL)
        plt.xlabel(ASSESSMENTTYPELBL)
        plt.ylabel(AVERAGESCORELBL)
        plt.grid(True)
        plt.savefig(path)
        plt.close()


    def scatterplot(self, path):
        # Plotting scatter plot between test and evaluation scores
        plt.scatter(self.file[EVALUATION], self.file[EXAM])
        plt.title(EVALUATIONVSEXAMSCORELBL)
        plt.xlabel(EVALUATIONLBL)
        plt.ylabel(EXAMLBL)
        plt.grid(True)
        plt.savefig(path)
        plt.close()


    def piechart(self, path):
        # Plotting the pie chart for score ranges
        self.percentagedistribution()
        labelslbl = [item + CATEGORYRATIONLBL for item in labels]
        plt.pie(self.analysetable.to_list(), labels=labelslbl, autopct=' %1.1f%%', startangle=90)
        plt.title(PERCENTAGEDISTRIBUTIONOFSCORELBL)
        plt.savefig(path)
        plt.close()


    def corrolation(self, path):
        # Correlation matrix [EVALUATION, ASSIGNMENT, EXAM]
        cor_matrix = self.file[[EVALUATION, ASSIGNMENT, EXAM]].corr()
        # Plotting the heatmap
        sns.heatmap(cor_matrix, annot=True, cmap='coolwarm')
        plt.title(CORRELATIONBETWEENASSESSMENTLBL)
        plt.savefig(path)
        plt.close()

