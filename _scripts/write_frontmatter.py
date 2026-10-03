# -*- coding: utf-8 -*-
"""Viet phan dau luan van + muc 4.10 vao DTL_Master_Thesis_Draft.docx"""
import docx
from docx.shared import Pt
from write_chapter3 import Writer

FILE = "DTL_Master_Thesis_Draft.docx"

ABBR = [
 ("ANOVA", "Analysis of Variance"),
 ("APG", "Advance Payment Guarantee"),
 ("BANK_REP", "Bank Reputation (independent construct X4)"),
 ("BG", "Payment Guarantee"),
 ("Big4", "The four largest state-owned commercial banks in Vietnam"),
 ("CCF", "Credit Conversion Factor (Basel framework)"),
 ("COLL_POLICY", "Collateral and Margin Flexibility (independent construct X7)"),
 ("COST_COMP", "Price Competitiveness (independent construct X1)"),
 ("DEC", "Selection Priority and Patronage Intention (dependent construct)"),
 ("DIGITAL_CONV", "Digital eFAST Convenience (independent construct X3)"),
 ("EFA", "Exploratory Factor Analysis"),
 ("eFAST", "VietinBank corporate digital banking platform"),
 ("FDI", "Foreign Direct Investment"),
 ("HSD", "Honestly Significant Difference (Tukey post-hoc test)"),
 ("ICC", "International Chamber of Commerce"),
 ("ISS", "International Institute of Social Studies, Erasmus University Rotterdam"),
 ("KMO", "Kaiser–Meyer–Olkin measure of sampling adequacy"),
 ("MDE", "Master's Programme in Development Economics"),
 ("NEU", "National Economics University"),
 ("NIM", "Net Interest Margin"),
 ("OLS", "Ordinary Least Squares"),
 ("PG", "Performance Guarantee"),
 ("PROC_SPEED", "Processing Speed (independent construct X2)"),
 ("RELATIONSHIP", "Relationship Banking and Limits (independent construct X5)"),
 ("RG", "Re-guarantee / Counter-guarantee"),
 ("RM", "Relationship Manager"),
 ("RQ", "Research Question"),
 ("SBV", "State Bank of Vietnam"),
 ("SERVQUAL", "Service Quality measurement model"),
 ("SME", "Small and Medium-sized Enterprise"),
 ("SOE", "State-Owned Enterprise"),
 ("STAFF_QUAL", "Staff Professionalism (independent construct X6)"),
 ("STP", "Straight-Through Processing"),
 ("TAM", "Technology Acceptance Model"),
 ("TG", "Tender Guarantee / Bid Bond"),
 ("URDG 758", "Uniform Rules for Demand Guarantees, ICC Publication No. 758"),
 ("VIF", "Variance Inflation Factor"),
 ("VND", "Vietnamese Dong"),
]


def settext(p, new, italic=False):
    if not p.runs:
        r = p.add_run(new); r.italic = italic; return
    p.runs[0].text = new
    p.runs[0].italic = italic
    for r in p.runs[1:]:
        r.text = ""


def main():
    doc = docx.Document(FILE)

    # ---------- 1. STATEMENT OF AUTHORSHIP: viet lai cho dung chuan
    for p in doc.paragraphs:
        if p.text.strip().startswith("I hereby declare that this Master"):
            settext(p,
                "I hereby declare that this Master's thesis entitled \"Factors Affecting Corporate Customers' "
                "Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry "
                "and Trade (VietinBank)\" is my own independent research work, carried out under the academic "
                "supervision of Dr. Hoang Thi Thuy Nga. All data reported in this thesis were collected by the "
                "author through the survey instrument described in Chapter 3, and all analyses were performed by "
                "the author. Any material drawn from the work of others is acknowledged in the text and listed "
                "in the references. This thesis has not been submitted, in whole or in part, for any other "
                "degree or qualification at this or any other institution.")
            print("Da viet lai Statement of Authorship")
            break

    # ---------- 2. ACKNOWLEDGEMENTS
    w = Writer(doc)
    w.at("ACKNOWLEDGEMENTS")
    w.para("This thesis would not have been completed without the support of a number of people and "
           "institutions, to whom I owe sincere thanks.")
    w.para("My deepest gratitude goes to my academic supervisor, Dr. Hoang Thi Thuy Nga, whose guidance shaped "
           "this study at every stage. Her comments on the initial research design led me to reduce and sharpen "
           "the model, to correct the measurement direction of the pricing construct, and to reconstruct the "
           "dependent variable so that it measures what the research question actually asks. The thesis is "
           "considerably stronger for that guidance, and any remaining shortcomings are entirely my own.")
    w.para("I am grateful to the faculty and administrative staff of the Vietnam–Netherlands Master's Programme "
           "in Development Economics at the National Economics University, and to the International Institute of "
           "Social Studies of Erasmus University Rotterdam, for the training in economics and quantitative "
           "methods on which this work rests.")
    w.para("I would also like to thank the officers of Vietnam Joint Stock Commercial Bank for Industry and "
           "Trade who assisted with the distribution of the survey, and the corporate clients who took the time "
           "to complete it. Their willingness to share considered evaluations of the guarantee service made the "
           "empirical part of this study possible.")
    w.para("Finally, I thank my family and my colleagues for their patience and encouragement throughout the "
           "period of study.")
    w.para("Hanoi, [tháng] 2026", after=2, first_line=0)
    w.para("Dang Tu Linh", first_line=0)
    print("Da viet Acknowledgements")

    # ---------- 3. LIST OF ABBREVIATIONS: cap nhat bang
    tb = doc.tables[0]
    while len(tb.rows) > 1:
        tb._element.remove(tb.rows[-1]._element)
    for k, v in ABBR:
        cells = tb.add_row().cells
        for i, txt in enumerate((k, v)):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(txt)
            r.font.name = 'Times New Roman'; r.font.size = Pt(11)
            if i == 0: r.bold = True
    print("Da cap nhat bang viet tat:", len(ABBR), "muc")

    doc.save(FILE)


if __name__ == "__main__":
    main()
