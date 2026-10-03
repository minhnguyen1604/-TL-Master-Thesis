# -*- coding: utf-8 -*-
"""Bo sung 7 tai lieu phuong phap va rai in-text citation cho Chuong 1 & 3."""
import docx, copy
from docx.text.paragraph import Paragraph

FILE = "DTL_Master_Thesis_Draft.docx"

# ---- tai lieu moi: (chuoi dau cua muc SE DUNG lam moc chen truoc, noi dung)
NEW_REFS = [
 ("Carletti, E.",
  "Brislin, R.W. (1970) 'Back-Translation for Cross-Cultural Research', Journal of Cross-Cultural "
  "Psychology, 1(3), pp. 185–216."),
 ("Davis, F.D.",
  "Cohen, J. (1988) Statistical Power Analysis for the Behavioral Sciences. 2nd edn. Hillsdale, NJ: "
  "Lawrence Erlbaum Associates."),
 ("Hair, J.F.",
  "Field, A. (2018) Discovering Statistics Using IBM SPSS Statistics. 5th edn. London: Sage Publications."),
 ("Oliver, R.L.",
  "Nunnally, J.C. and Bernstein, I.H. (1994) Psychometric Theory. 3rd edn. New York: McGraw-Hill."),
 ("Ramakrishnan, R.T.S.",
  "Podsakoff, P.M., MacKenzie, S.B., Lee, J.-Y. and Podsakoff, N.P. (2003) 'Common Method Biases in "
  "Behavioral Research: A Critical Review of the Literature and Recommended Remedies', Journal of Applied "
  "Psychology, 88(5), pp. 879–903."),
 ("Turnbull, P.W.",
  "Tabachnick, B.G. and Fidell, L.S. (2019) Using Multivariate Statistics. 7th edn. Boston, MA: Pearson."),
 ("Vietnamese Government",
  "VietinBank (n.d.) Annual Report and Audited Consolidated Financial Statements. Hanoi: Vietnam Joint Stock "
  "Commercial Bank for Industry and Trade. [Cần bổ sung năm báo cáo cụ thể khi trích dẫn số liệu tại mục 1.2]"),
]

# ---- in-text citation: (chuoi tim, chuoi thay)
EDITS = [
 # ---------- 1.2
 ("operating a nationwide corporate banking network of 155 domestic branches together with the VietinBank "
  "eFAST digital platform for corporate clients.",
  "operating a nationwide corporate banking network of 155 domestic branches together with the VietinBank "
  "eFAST digital platform for corporate clients (VietinBank, n.d.)."),
 ("A second dimension of the problem concerns segmentation. A uniform competitive posture across the "
  "entire corporate base is unlikely to be efficient if, as the international literature suggests, "
  "selection criteria vary with firm characteristics.",
  "A second dimension of the problem concerns segmentation. A uniform competitive posture across the "
  "entire corporate base is unlikely to be efficient if, as the international literature suggests "
  "(Turnbull and Gibbs, 1989; Kaur et al., 2021), selection criteria vary with firm characteristics."),
 # ---------- 3.2.2
 ("Three procedures protect the content validity of the instrument.",
  "Three procedures protect the content validity of the instrument, following the guidance of Hair et al. "
  "(2019) on scale development."),
 ("translated, with a back-translation into English by a second translator to confirm that no "
  "meaning was lost; discrepancies were reconciled before finalisation.",
  "translated, with a back-translation into English by a second translator to confirm that no "
  "meaning was lost, following the procedure of Brislin (1970); discrepancies were reconciled before "
  "finalisation."),
 # ---------- 3.3.1
 ("The study therefore applies stratified convenience sampling.",
  "The study therefore applies stratified convenience sampling, an approach widely used where the "
  "population list is not accessible to the researcher (Hair et al., 2019)."),
 ("This procedure is not probabilistic, and the limitations that follow are stated openly.",
  "This procedure is not probabilistic, and the limitations that follow are stated openly (Hair et al., "
  "2019)."),
 # ---------- 3.3.3
 ("Administration through the bank's own relationship channel carries a recognised risk of social "
  "desirability bias,",
  "Administration through the bank's own relationship channel carries a recognised risk of social "
  "desirability bias, a form of common method bias documented by Podsakoff et al. (2003),"),
 ("Three measures limit this risk:",
  "Three of the procedural remedies recommended by Podsakoff et al. (2003) are applied to limit this risk:"),
 # ---------- 3.4.1
 ("Skewness and kurtosis are examined against the conventional tolerance of plus or minus one for "
  "skewness and plus or minus three for kurtosis;",
  "Skewness and kurtosis are examined against the conventional tolerance of plus or minus one for "
  "skewness and plus or minus three for kurtosis (Hair et al., 2019; Tabachnick and Fidell, 2019);"),
 ("Questionnaires missing any Part II indicator are discarded rather than imputed, since listwise "
  "deletion is unproblematic at the planned sample size.",
  "Questionnaires missing any Part II indicator are discarded rather than imputed, since listwise "
  "deletion is unproblematic at the planned sample size (Tabachnick and Fidell, 2019)."),
 # ---------- 3.4.2
 ("Following Hair et al. (2019), a construct is accepted where alpha reaches 0.70.",
  "Following Nunnally and Bernstein (1994) and Hair et al. (2019), a construct is accepted where alpha "
  "reaches 0.70."),
 ("The corrected item-total correlation must reach 0.30;",
  "The corrected item-total correlation must reach 0.30 (Field, 2018);"),
 # ---------- 3.4.3
 ("Principal component extraction with Varimax orthogonal rotation is applied.",
  "Principal component extraction with Varimax orthogonal rotation is applied (Hair et al., 2019; "
  "Field, 2018)."),
 ("The criteria applied are set out in Table 3.2.",
  "The criteria applied, drawn from Hair et al. (2019) and Field (2018), are set out in Table 3.2."),
 # ---------- 3.4.4
 ("Regression requires one value per construct per enterprise. Two approaches are available.",
  "Regression requires one value per construct per enterprise. Two approaches are available "
  "(Hair et al., 2019)."),
 # ---------- 3.4.5
 ("This study applies a conservative threshold of VIF below 3.0, corresponding to tolerance above 0.33, "
  "rather than the more permissive value of 10 found in some textbooks.",
  "This study applies a conservative threshold of VIF below 3.0, corresponding to tolerance above 0.33, "
  "rather than the more permissive value of 10 reported in Hair et al. (2019) and Field (2018)."),
 ("a coefficient of 0.80 or above between two predictors signals that they may not be empirically distinct.",
  "a coefficient of 0.80 or above between two predictors signals that they may not be empirically distinct "
  "(Tabachnick and Fidell, 2019)."),
 # ---------- 3.4.6
 ("Four assumptions underlying OLS are tested and the results reported in Section 4.5.1.",
  "Four assumptions underlying OLS are tested, following Tabachnick and Fidell (2019) and Field (2018), "
  "and the results reported in Section 4.5.1."),
 ("Independence of errors is checked through the Durbin–Watson statistic, with values between 1.5 and "
  "2.5 considered acceptable.",
  "Independence of errors is checked through the Durbin–Watson statistic, with values between 1.5 and "
  "2.5 considered acceptable (Field, 2018)."),
 # ---------- 3.4.7
 ("Levene's test is applied first to verify homogeneity of variance; where it is rejected, Welch's "
  "adjusted F-statistic is reported in place of the conventional one.",
  "Levene's test is applied first to verify homogeneity of variance; where it is rejected, Welch's "
  "adjusted F-statistic is reported in place of the conventional one (Field, 2018)."),
 ("Eta squared is reported for analysis of variance and Cohen's d for the t-test.",
  "Eta squared is reported for analysis of variance and Cohen's d for the t-test, interpreted against the "
  "conventions of Cohen (1988)."),
]


def settext(p, new):
    if not p.runs:
        p.add_run(new); return
    p.runs[0].text = new
    for r in p.runs[1:]:
        r.text = ""


def main():
    d = docx.Document(FILE)

    # --- 1. chen tai lieu moi
    added = 0
    for anchor_prefix, ref in NEW_REFS:
        target = None
        for p in d.paragraphs:
            if p.text.strip().startswith(anchor_prefix):
                target = p; break
        if target is None:
            print("  !! khong tim thay moc:", anchor_prefix); continue
        el = copy.deepcopy(target._element)
        target._element.addprevious(el)
        settext(Paragraph(el, target._parent), ref)
        added += 1
    print("Da chen", added, "tai lieu phuong phap vao danh muc")

    # --- 2. rai in-text citation
    hit = miss = 0
    for old, new in EDITS:
        done = False
        for p in d.paragraphs:
            if old in p.text:
                settext(p, p.text.replace(old, new)); done = True; break
        if done: hit += 1
        else:
            miss += 1; print("  !! khong khop:", old[:70])
    print("In-text citation: thay duoc", hit, "/", hit + miss)
    d.save(FILE)


if __name__ == "__main__":
    main()
