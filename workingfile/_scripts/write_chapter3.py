# -*- coding: utf-8 -*-
"""Viet noi dung Chuong 3 vao DTL_Master_Thesis_Draft.docx"""
import docx, copy
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

FILE = "DTL_Master_Thesis_Draft.docx"
NAVY = RGBColor(0x00, 0x33, 0x66)
GREY = "D9D9D9"


def shade(cell, f):
    tcPr = cell._element.get_or_add_tcPr()
    s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), f); tcPr.append(s)


class Writer:
    """Chen noi dung ngay sau mot de muc trong tai lieu."""
    def __init__(self, doc):
        self.doc = doc
        self.cursor = None

    def at(self, heading_text):
        for p in self.doc.paragraphs:
            if p.text.strip() == heading_text:
                self.cursor = p
                return self
        raise ValueError("Khong tim thay de muc: " + heading_text)

    def _new_par(self):
        el = OxmlElement('w:p')
        self.cursor._element.addnext(el)
        p = Paragraph(el, self.cursor._parent)
        self.cursor = p
        return p

    def para(self, text, italic=False, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             size=12, after=8, indent=None, first_line=0.25):
        p = self._new_par()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_after = Pt(after); pf.line_spacing = 1.5
        if indent is not None: pf.left_indent = Inches(indent)
        if first_line: pf.first_line_indent = Inches(first_line)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'; r.font.size = Pt(size)
        r.italic = italic; r.bold = bold
        return p

    def bullet(self, text, size=12):
        p = self._new_par()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_after = Pt(5); pf.line_spacing = 1.4
        pf.left_indent = Inches(0.4); pf.first_line_indent = Inches(-0.2)
        r = p.add_run("– " + text)
        r.font.name = 'Times New Roman'; r.font.size = Pt(size)
        return p

    def caption(self, text):
        return self.para(text, italic=False, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         size=11, after=4, first_line=0)

    def source(self, text):
        return self.para(text, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         size=10, after=10, first_line=0)

    def equation(self, text):
        return self.para(text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         size=12, after=10, first_line=0)

    def table(self, rows, widths, font=9.5):
        tb = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(rows):
            for k, v in enumerate(row):
                cell = tb.rows[i].cells[k]; cell.width = Inches(widths[k]); cell.text = ""
                p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                if i == 0 or k >= len(row) - 1:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(v)
                r.font.name = 'Times New Roman'; r.font.size = Pt(font)
                if i == 0:
                    r.bold = True; shade(cell, GREY)
        # di chuyen bang ve dung vi tri con tro
        self.cursor._element.addnext(tb._element)
        self.cursor = Paragraph(tb._element.getnext(), self.cursor._parent) \
            if tb._element.getnext() is not None and tb._element.getnext().tag.endswith('}p') else self.cursor
        # chen mot doan trong sau bang de tiep tuc
        el = OxmlElement('w:p')
        tb._element.addnext(el)
        self.cursor = Paragraph(el, self.doc)
        return tb


ITEMS = [
 ("COST_COMP", "Price Competitiveness", "COMP1",
  "The guarantee issuance fee charged by VietinBank is reasonable relative to the service quality received."),
 ("", "", "COMP2", "VietinBank's guarantee fee schedule is more competitive than those of other banks the enterprise has examined."),
 ("", "", "COMP3", "VietinBank applies flexible fee incentives and discounts for regular corporate clients."),
 ("PROC_SPEED", "Processing Speed", "SPEED1", "Appraisal and approval of the guarantee limit at VietinBank is fast."),
 ("", "", "SPEED2", "The application procedure for a guarantee at VietinBank is simple and free of unnecessary paperwork."),
 ("", "", "SPEED3", "The time from submission of a complete file to receipt of the letter of guarantee meets contract deadlines."),
 ("DIGITAL_CONV", "Digital eFAST Convenience", "DIGI1", "The enterprise can submit guarantee applications online 24/7 through the VietinBank eFAST platform."),
 ("", "", "DIGI2", "VietinBank issues digitally signed electronic letters of guarantee promptly, without a paper original."),
 ("", "", "DIGI3", "Online look-up, authentication and status tracking of letters of guarantee is convenient and reliable."),
 ("BANK_REP", "Bank Reputation", "REPU1", "VietinBank's reputation and brand rank among the leaders of the Vietnamese banking market."),
 ("", "", "REPU2", "Letters of guarantee issued by VietinBank are accepted by project owners and procuring entities more readily than those of many other banks."),
 ("", "", "REPU3", "VietinBank's financial strength adds credibility to the enterprise when bidding and signing contracts."),
 ("RELATIONSHIP", "Relationship & Limits", "RELA1", "A long-standing credit relationship with VietinBank makes it easier for the enterprise to obtain a guarantee."),
 ("", "", "RELA2", "Using several other VietinBank services (accounts, payments, payroll, funded credit) brings advantages for guarantee transactions."),
 ("", "", "RELA3", "VietinBank grants and adjusts guarantee limits flexibly according to the enterprise's actual needs."),
 ("STAFF_QUAL", "Staff Professionalism", "STAFF1", "VietinBank relationship managers have solid technical expertise in guarantee operations."),
 ("", "", "STAFF2", "VietinBank officers can advise on guarantee wording consistent with the Law on Bidding, Circular 61/2024/TT-NHNN and URDG 758."),
 ("", "", "STAFF3", "VietinBank officers provide timely and attentive support when guarantee-related problems arise."),
 ("COLL_POLICY", "Collateral & Margin Flexibility", "COLL1", "VietinBank applies flexible cash margin ratios, including reduction or waiver for qualifying clients."),
 ("", "", "COLL2", "VietinBank accepts a diverse range of collateral (real estate, machinery and equipment, contract receivables)."),
 ("", "", "COLL3", "Valuation and collateral registration procedures at VietinBank are quick and convenient."),
 ("DEC", "Selection Priority & Patronage Intention", "DEC1", "When guarantee needs arise, the enterprise prioritises VietinBank ahead of other banks."),
 ("", "", "DEC2", "The enterprise allocates the majority of its guarantee value and volume to VietinBank rather than to other banks."),
 ("", "", "DEC3", "For upcoming tenders and contracts, the enterprise intends to continue with VietinBank rather than switch."),
 ("", "", "DEC4", "The enterprise would recommend VietinBank to partners and joint-venture contractors as the best issuing bank."),
]


def write(w):
    # ============================ 3.1
    w.at("3.1. Overall Research Design & Analytical Process")
    w.para("This study adopts a quantitative research design grounded in the positivist paradigm. The logic of "
           "enquiry is deductive: hypotheses H1 to H8 are derived from the theoretical framework established in "
           "Chapter 2 and are then subjected to empirical testing against primary survey data. A quantitative "
           "approach is appropriate because the research questions concern not merely which attributes corporate "
           "customers consider, but the direction, statistical significance and relative magnitude of each "
           "attribute's association with selection priority. Answering RQ2 in particular requires a ranking of "
           "effect sizes that only a multivariate estimation procedure can deliver.")
    w.para("The design is cross-sectional. Data are collected at a single point in time from a large sample of "
           "corporate enterprises, rather than tracking the same enterprises over successive periods. This choice "
           "reflects the practical constraints of a master's thesis and is standard in the corporate bank "
           "selection literature (Turnbull and Gibbs, 1989; Narteh, 2013; Kaur et al., 2021; Zelie, 2023). Its "
           "principal limitation, namely the inability to establish temporal precedence and therefore strict "
           "causality, is acknowledged in Section 4.10; the relationships reported in this thesis are "
           "associational rather than causal.")
    w.para("The unit of analysis is the enterprise, not the individual respondent. Each questionnaire is "
           "completed by one officer authorised to speak for the enterprise on guarantee matters, and the "
           "responses are interpreted as organisational rather than personal evaluations. This is consistent with "
           "the organisational buying behaviour perspective adopted in Chapter 2.")
    w.para("Following the analytical sequence approved by the supervisor, the empirical work proceeds through "
           "seven steps, each serving a distinct methodological purpose:")
    for t in [
        "Step 1 – Descriptive statistics. Profiling the sample and screening the data for quality problems before "
        "any inferential procedure is applied.",
        "Step 2 – Reliability testing (Cronbach's Alpha). Verifying that the four indicators of each construct "
        "measure a common underlying dimension consistently.",
        "Step 3 – Exploratory Factor Analysis. Establishing that the observed indicators group into the factor "
        "structure the conceptual framework predicts, and that constructs are empirically distinguishable.",
        "Step 4 – Factor score extraction. Converting the retained indicators into a single representative "
        "variable per construct for use in regression.",
        "Step 5 – Pearson correlation analysis. Examining the bivariate association between each independent "
        "construct and the dependent construct, and detecting early signs of excessive inter-construct correlation.",
        "Step 6 – Multicollinearity diagnostics. Confirming through variance inflation factors that the "
        "independent constructs do not overlap to a degree that would destabilise the coefficient estimates.",
        "Step 7 – OLS regression and sub-group difference testing. Estimating the magnitude and significance of "
        "each construct's association with selection priority, and testing whether that association differs "
        "across firm characteristics."]:
        w.bullet(t)
    w.para("Steps 2 and 3 function jointly as the measurement validation stage: no construct enters the "
           "regression model until its indicators have demonstrated both internal consistency and convergent and "
           "discriminant validity. All computations are performed in Python, using the pandas library for data "
           "management, factor_analyzer for reliability and factor analysis, statsmodels for regression and "
           "diagnostics, and pingouin for group difference tests. Full outputs are reproduced in Appendices 1 "
           "to 4.")

    # ============================ 3.2.1
    w.at("3.2.1. Operationalization of Variables")
    w.para("All eight constructs in the research model are latent variables: they cannot be observed directly and "
           "must be inferred from responses to a set of observed indicators. Each construct is modelled as "
           "reflective, meaning that the latent construct is understood to cause the observed responses, so that "
           "the indicators of a given construct should correlate strongly with one another and be "
           "interchangeable manifestations of the same underlying dimension.")
    w.para("Each construct is measured by four indicators, including the dependent construct, which "
           "is measured by four. Three indicators is the minimum required for a construct to be identified in "
           "factor analysis and is adopted deliberately here to keep the instrument short. This decision responds "
           "directly to the supervisor's guidance on model parsimony and to the practical requirement that a "
           "questionnaire aimed at senior corporate officers can be completed in a few minutes; over-long "
           "instruments invite non-response and straight-lining, both of which degrade data quality more "
           "seriously than the marginal precision gained from additional items.")
    w.para("Every indicator is scored on a five-point Likert scale anchored at 1 (strongly disagree) and 5 "
           "(strongly agree). Following standard practice in service quality and bank selection research, the "
           "resulting summated scales are treated as interval-level data for the purposes of parametric analysis. "
           "Table 3.1 presents the full operationalisation.")
    w.caption("Table 3.1: Operationalisation of constructs and measurement indicators")
    rows = [["Construct", "Code", "Measurement indicator"]]
    for c, name, code, text in ITEMS:
        rows.append([(c + "\n(" + name + ")") if c else "", code, text])
    w.table(rows, [1.35, 0.65, 4.2], font=9)
    w.source("Source: compiled by the author from the theoretical framework in Chapter 2.")
    w.para("The seven independent constructs are each anchored in the theory identified in Section 2.2. Price "
           "Competitiveness and Collateral Policy derive from credit risk pricing theory (Merton, 1974; Stiglitz "
           "and Weiss, 1981), under which guarantee fees and margin requirements are the price of a contingent "
           "credit exposure. Processing Speed and Staff Professionalism correspond to the responsiveness and "
           "assurance dimensions of SERVQUAL (Parasuraman, Zeithaml and Berry, 1988). Bank Reputation follows "
           "from delegated monitoring theory (Diamond, 1984; Ramakrishnan and Thakor, 1984), whereby the "
           "issuing bank's standing is itself the signal the beneficiary relies upon. Relationship and Limits "
           "follows relationship banking theory (Boot, 2000; Berger and Udell, 1995). Digital eFAST Convenience "
           "follows the Technology Acceptance Model (Davis, 1989; Venkatesh et al., 2003).")
    w.para("The direction of measurement deserves particular attention for the fee construct. The three "
           "indicators COMP1 to COMP3 do not measure the absolute level of fees; they measure the enterprise's "
           "evaluation of how reasonable and competitive those fees are. A higher score therefore denotes a more "
           "favourable perception, and the hypothesised association with selection priority is positive. Labelling "
           "the construct Price Competitiveness rather than Cost makes this explicit and removes the sign "
           "inconsistency that would otherwise arise from treating a favourably-worded evaluation as a cost "
           "variable.")
    w.para("The dependent construct requires a further clarification. Selection in the strict econometric sense "
           "is a discrete act, and a study observing actual switching between banks would call for a logit or "
           "probit specification. This thesis does not observe switching; it measures the degree of priority an "
           "enterprise assigns to VietinBank relative to competing banks, expressed on a continuous multi-item "
           "attitudinal scale. The four indicators are adapted from the behavioural consequences battery of "
           "Zeithaml, Berry and Parasuraman (1996), which captures preference, allocation, continuation intention "
           "and advocacy. Oliver's (1999) four-stage loyalty framework is invoked at the conceptual level to "
           "justify treating patronage as a multi-dimensional attitudinal construct rather than a single "
           "question, but no one-to-one correspondence between the four indicators and the four stages is "
           "claimed: advocacy does not constitute a stage in that framework, and action loyalty cannot be "
           "observed in a cross-sectional design. Because the construct is continuous rather than binary, "
           "ordinary least squares estimation is the appropriate technique.")

    # ============================ 3.2.2
    w.at("3.2.2. Mapping Scales with the Official 32-Item Survey Questionnaire")
    w.para("The survey instrument is organised in two parts. Part I collects six items of classification "
           "information: ownership type, annual revenue band, operating tenure, the guarantee product the "
           "enterprise uses most frequently, the number of banks at which it currently obtains guarantees, and "
           "the position of the responding officer. Part II contains the thirty-two Likert indicators listed in "
           "Table 3.1, presented in eight blocks corresponding to the eight constructs.")
    w.para("The instrument is deliberately constructed so that the mapping between questionnaire and model is "
           "exhaustive in both directions. Every indicator in the questionnaire belongs to a construct that "
           "appears in the regression equation, and every construct in the equation is measured by indicators "
           "that appear in the questionnaire. There are no residual items collected but unused, and no construct "
           "specified without a corresponding measure. The indicator codes themselves carry the construct name, "
           "so that the link between the raw data file and the estimated model is unambiguous at every stage of "
           "the analysis. This design responds directly to the supervisor's instruction that the measurement "
           "scales must correspond to the questionnaire actually administered rather than being constructed "
           "independently and mapped onto the data after the fact.")
    w.para("Three procedures protect the content validity of the instrument. First, each indicator is adapted "
           "from a published scale and reworded to the bank guarantee context rather than invented; the "
           "provenance of each construct is set out in Section 2.2 and summarised in Table 3.1. Second, the "
           "questionnaire is administered in Vietnamese, having been developed in English from the source scales "
           "and translated, with a back-translation into English by a second translator to confirm that no "
           "meaning was lost; discrepancies were reconciled before finalisation. Third, the draft instrument is "
           "reviewed by experienced guarantee officers and by a small number of corporate clients in a pilot "
           "round, whose comments are used to remove technical ambiguity and to confirm that the wording is "
           "intelligible to the intended respondents. The pilot responses are used only for instrument "
           "refinement and are excluded from the main dataset.")
    w.para("The dependent construct block is introduced by an explicit comparative instruction asking the "
           "respondent to answer by reference to the other banks the enterprise has used or considered for "
           "guarantee services. Without this framing, agreement with a statement such as \"the enterprise "
           "prioritises VietinBank\" could be read as a general expression of satisfaction. The instruction "
           "anchors the entire block in the competitive comparison that the research question requires.")

    # ============================ 3.3.1
    w.at("3.3.1. Target Population & Sampling Method")
    w.para("The target population comprises enterprises that are current users of bank guarantee services at "
           "VietinBank. Three boundary conditions define membership. The enterprise must be a legal entity rather "
           "than a household or individual business, since the corporate guarantee product is not offered to the "
           "latter. It must have had at least one guarantee issued or outstanding at VietinBank within the twelve "
           "months preceding the survey, so that the respondent is evaluating recent rather than remembered "
           "experience. And the respondent must be an officer with direct responsibility for guarantee "
           "transactions, so that the evaluation reflects informed organisational judgement.")
    w.para("The sampling frame is the population of corporate guarantee clients across VietinBank's 155 domestic "
           "branches. Because a complete list of qualifying enterprises cannot be released to an external "
           "researcher for confidentiality reasons, simple random selection from the frame is not feasible. The "
           "study therefore applies stratified convenience sampling. The branch network is first stratified by "
           "geographic region, and a quota of responses is allocated to each region in proportion to its share of "
           "the bank's corporate guarantee balance, so that the regional composition of the sample approximates "
           "that of the population. Within each stratum, questionnaires are distributed to whichever qualifying "
           "clients are accessible through the branch relationship managers.")
    w.para("This procedure is not probabilistic, and the limitations that follow are stated openly. Selection "
           "probabilities are unknown, so sampling error cannot be estimated and inference to the population "
           "rests on the assumption that respondents do not differ systematically from non-respondents on the "
           "constructs measured. Stratification by region mitigates but does not eliminate this risk. More "
           "fundamentally, the frame contains only enterprises that already use VietinBank; enterprises that "
           "considered the bank and chose a competitor, or that used it and left, are not represented. The "
           "consequence, discussed further in Section 4.10, is that the findings describe the determinants of "
           "selection priority and patronage within the existing client base rather than the determinants of "
           "first-time bank choice across the market as a whole. The research scope stated in Section 1.5 is "
           "defined accordingly.")

    # ============================ 3.3.2
    w.at("3.3.2. Sample Size Determination")
    w.para("Three conventional criteria govern the minimum sample size for the planned analysis. For exploratory "
           "factor analysis, Hair et al. (2019) recommend a minimum ratio of five observations per measured "
           "indicator, which for thirty-two indicators yields 160 responses, and regard a ratio of ten to one, "
           "or 250 responses, as comfortable. For multiple regression, the widely applied rule of 50 + 8k, where "
           "k is the number of predictors, yields a minimum of 106 responses for seven predictors. For sub-group "
           "comparison, the binding constraint is not the total but the size of the smallest cell: one-way "
           "analysis of variance requires roughly thirty observations in each group for the test to be robust to "
           "moderate departures from normality, and the ownership and product classifications each define five "
           "groups.")
    w.para("The target sample of 800 valid responses exceeds all three thresholds by a wide margin. The binding "
           "criterion is the third: with 800 responses distributed across five ownership categories, even a "
           "category holding only five per cent of the sample would contain forty enterprises, sufficient for the "
           "group comparisons required by Objective 3. The target also provides headroom for attrition, since "
           "questionnaires that are incomplete, internally inconsistent or straight-lined must be discarded "
           "before analysis. Assuming a usable rate of around seventy per cent, approximately 1,150 "
           "questionnaires need to be distributed to yield 800 valid responses.")
    w.para("A sample of this size carries one methodological consequence that must be anticipated rather than "
           "discovered. With n = 800, the standard errors of the estimated coefficients are small, and "
           "coefficients of negligible practical magnitude may nonetheless attain conventional significance "
           "levels. Statistical significance alone therefore conveys limited information in this setting. The "
           "interpretation in Chapter 4 accordingly rests on standardised beta coefficients, which permit the "
           "seven constructs to be ranked by strength of association, and on the adjusted coefficient of "
           "determination, which indicates how much of the variation in selection priority the model explains. "
           "This emphasis on practical alongside statistical significance is what makes the results usable as a "
           "basis for managerial prioritisation.")

    # ============================ 3.3.3
    w.at("3.3.3. Survey Administration across 155 VietinBank Branches")
    w.para("Questionnaires are distributed through the corporate relationship managers of VietinBank branches, "
           "who are in regular contact with the qualifying enterprises and can identify officers competent to "
           "respond. Distribution takes place in two modes: a printed form handed over during scheduled client "
           "meetings, and an identical electronic form sent by email to clients who prefer to respond remotely. "
           "Each returned questionnaire carries a branch code, which permits the regional quotas to be monitored "
           "during fieldwork and allows the achieved composition of the sample to be compared with the intended "
           "stratification.")
    w.para("Participation is voluntary and responses are anonymous. The covering letter states the purpose of "
           "the survey, confirms that responses are used solely for aggregate analysis, and gives an undertaking "
           "that no individual enterprise or respondent is identified at any stage. No commercially sensitive "
           "information about guarantee balances, pricing or credit terms is requested; all items are perceptual, "
           "which reduces both the sensitivity of the exercise and the burden on the respondent.")
    w.para("Administration through the bank's own relationship channel carries a recognised risk of social "
           "desirability bias, since respondents may hesitate to record unfavourable evaluations to an officer "
           "with whom they maintain a working relationship. Three measures limit this risk: the questionnaire is "
           "returned in a sealed envelope or submitted directly through the electronic form rather than handed "
           "back to the relationship manager; the covering letter emphasises anonymity; and no identifying "
           "information beyond the classification items in Part I is collected. Residual bias is assessed at the "
           "descriptive stage by inspecting the item means and standard deviations, since a distribution "
           "compressed towards the upper end of the scale with unusually small dispersion is the characteristic "
           "signature of the problem. Where present, it is reported and taken into account in interpretation; it "
           "does not invalidate the regression results, which depend on variation between respondents rather than "
           "on the absolute level of the means.")
    w.para("[Cần bổ sung: căn cứ chấp thuận của VietinBank cho việc triển khai khảo sát qua mạng lưới chi nhánh "
           "— số văn bản, đơn vị ban hành, thời gian hiệu lực. Đoạn này viết sau khi có văn bản.]",
           italic=True)

    # ============================ 3.4.1
    w.at("3.4.1. Descriptive Statistics")
    w.para("Descriptive analysis serves two purposes: to establish the composition of the achieved sample, and "
           "to screen the data before inferential procedures are applied. The sample profile is reported as "
           "frequency distributions across the five classification variables of Part I, and the achieved "
           "distribution is compared with the intended regional stratification to document any imbalance.")
    w.para("For each of the thirty-two indicators the analysis reports the mean, standard deviation, minimum, "
           "maximum, skewness and kurtosis. Skewness and kurtosis are examined against the conventional "
           "tolerance of plus or minus one for skewness and plus or minus three for kurtosis; indicators falling "
           "outside these bounds are noted, since severe non-normality would affect the maximum likelihood "
           "extraction option in factor analysis and the inference procedures in regression.")
    w.para("Data screening removes three categories of response before analysis. Questionnaires missing any "
           "Part II indicator are discarded rather than imputed, since listwise deletion is unproblematic at the "
           "planned sample size. Straight-lined responses, defined as a questionnaire recording an identical "
           "value for every indicator across all eight blocks, are removed as non-informative. Internally "
           "inconsistent responses, identified through the small number of conceptually opposed item pairs, are "
           "reviewed individually. The number of responses excluded under each rule is reported so that the "
           "path from distributed questionnaires to the analytical sample is fully traceable.")

    # ============================ 3.4.2
    w.at("3.4.2. Scale Reliability Testing (Cronbach’s Alpha)")
    w.para("Reliability testing establishes whether the indicators assigned to each construct measure that "
           "construct consistently. Cronbach's alpha is computed separately for each of the eight blocks. "
           "Following Hair et al. (2019), a construct is accepted where alpha reaches 0.70. This threshold is "
           "set above the 0.60 sometimes accepted in exploratory work because the constructs here are adapted "
           "from established scales rather than newly developed, and because a three-indicator construct that "
           "cannot reach 0.70 is unlikely to behave stably in subsequent analysis.")
    w.para("Two supplementary statistics are examined for every indicator. The corrected item-total correlation "
           "must reach 0.30; an indicator falling below this contributes little to the construct and is a "
           "candidate for removal. The alpha-if-item-deleted statistic identifies indicators whose removal would "
           "raise the construct's alpha. Because each construct carries four indicators and three are needed for comfortable identification, the "
           "minimum for identification, removal is undertaken only where an indicator is clearly deficient on "
           "both criteria, and any such decision is reported explicitly with its justification rather than "
           "applied silently.")

    # ============================ 3.4.3
    w.at("3.4.3. Exploratory Factor Analysis (EFA)")
    w.para("Factor analysis examines whether the thirty-two indicators group empirically into the eight "
           "constructs the conceptual framework specifies. The procedure is described as exploratory in the "
           "sense that the factor structure is allowed to emerge from the correlation matrix rather than being "
           "imposed; the expected structure is nonetheless known in advance, so the analysis functions in "
           "practice as a test of the measurement model. Where the emergent structure departs from the expected "
           "one, the departure is reported and interpreted rather than suppressed.")
    w.para("Analysis is conducted separately for the twenty-eight independent indicators and the four dependent "
           "indicators. Running them jointly would allow the dependent construct to influence the factor "
           "solution of the predictors, which is methodologically undesirable in a model where one construct is "
           "the outcome of the others.")
    w.para("Principal component extraction with Varimax orthogonal rotation is applied. Orthogonal rotation is "
           "chosen because the subsequent regression requires factor scores that are as far as possible free of "
           "mutual correlation; an oblique rotation permitting correlated factors would aggravate the "
           "multicollinearity examined in Section 3.4.5. The criteria applied are set out in Table 3.2.")
    w.caption("Table 3.2: Criteria applied in exploratory factor analysis")
    w.table([
        ["Criterion", "Threshold", "Purpose"],
        ["Kaiser–Meyer–Olkin measure", "KMO ≥ 0.50; 0.80 or above desirable",
         "Confirms the correlation matrix is suitable for factor extraction"],
        ["Bartlett's test of sphericity", "p < 0.05",
         "Rejects the hypothesis that the correlation matrix is an identity matrix"],
        ["Eigenvalue", "≥ 1.00", "Determines the number of factors retained"],
        ["Cumulative variance explained", "≥ 50%", "Confirms the retained factors capture most of the shared variance"],
        ["Factor loading", "≥ 0.50", "Confirms each indicator loads substantively on its factor"],
        ["Cross-loading gap", "≥ 0.30 between primary and next-highest loading",
         "Confirms each indicator belongs unambiguously to one construct (discriminant validity)"],
    ], [1.7, 1.95, 2.55], font=9.5)
    w.source("Source: compiled by the author following Hair et al. (2019).")
    w.para("The expected outcome is a seven-factor solution for the independent indicators, with the three "
           "indicators of each construct loading together, and a single-factor solution for the dependent "
           "indicators. Two departures are anticipated as possibilities and their treatment is specified in "
           "advance. Should indicators from two constructs converge on a single factor, the constructs are "
           "conceptually adjacent and the merged factor is retained and renamed, with the theoretical "
           "implications discussed rather than the result forced back into the original structure. Should an "
           "indicator fail to load at 0.50 on any factor, or load on two factors with a gap below 0.30, it is "
           "removed and the analysis re-run, with the removal reported.")

    # ============================ 3.4.4
    w.at("3.4.4. Factor Scores Extraction Method")
    w.para("Regression requires one value per construct per enterprise. Two approaches are available. Regression "
           "factor scores weight each indicator by its loading and produce standardised variables that are "
           "orthogonal by construction. Summated mean scores take the unweighted arithmetic mean of the "
           "indicators retained after factor analysis.")
    w.para("This study uses summated mean scores. Two considerations favour them here. Interpretation is direct: "
           "a construct score of 4.2 is immediately readable on the original five-point scale, whereas a "
           "standardised factor score has meaning only relative to the sample and cannot be communicated to bank "
           "management without translation. Comparability is preserved: summated scores are computed identically "
           "across sub-samples, so the group comparisons in Section 3.4.7 and the descriptive statistics in "
           "Section 4.1 rest on the same metric, which is not guaranteed when factor scores are re-estimated "
           "within each group. The cost of this choice is that indicators contribute equally regardless of "
           "loading; given that all retained indicators must load at 0.50 or above, the loss of precision is "
           "small and is accepted in exchange for interpretability.")

    # ============================ 3.4.5
    w.at("3.4.5. Pearson Correlation Analysis & Multicollinearity Diagnostics (VIF)")
    w.para("The Pearson correlation matrix is computed across all eight construct scores. It serves two ends. "
           "Each independent construct should correlate significantly with the dependent construct, providing "
           "preliminary evidence consistent with the hypotheses before the multivariate model is estimated. And "
           "correlations among the independent constructs should remain moderate; a coefficient of 0.80 or above "
           "between two predictors signals that they may not be empirically distinct.")
    w.para("Multicollinearity is then assessed formally through variance inflation factors. This study applies a "
           "conservative threshold of VIF below 3.0, corresponding to tolerance above 0.33, rather than the more "
           "permissive value of 10 found in some textbooks. The stricter criterion is warranted by the structure "
           "of the model. In corporate banking, relationship depth, institutional reputation and collateral "
           "flexibility are empirically entangled: an enterprise with a long credit history typically also "
           "receives more accommodating margin terms and holds a favourable view of the bank's standing. Under "
           "these conditions a VIF between 3 and 10 would already destabilise the relative ranking of the "
           "standardised coefficients, and it is precisely that ranking on which the managerial conclusions of "
           "this thesis depend.")
    w.para("Should any construct exceed the threshold, three responses are available and the choice among them "
           "is reported: combining the affected constructs where factor analysis also indicates convergence, "
           "removing the construct with the weaker theoretical justification, or retaining the specification "
           "while explicitly qualifying the interpretation of the affected coefficients.")

    # ============================ 3.4.6
    w.at("3.4.6. Multiple Linear Regression Model Specification (OLS)")
    w.para("The hypotheses are tested through a multiple linear regression estimated by ordinary least squares. "
           "The specification is:")
    w.equation("DEC = β₀ + β₁·COST_COMP + β₂·PROC_SPEED + β₃·DIGITAL_CONV + β₄·BANK_REP")
    w.equation("+ β₅·RELATIONSHIP + β₆·STAFF_QUAL + β₇·COLL_POLICY + ε          (3.1)")
    w.para("where DEC is the summated selection priority score, the seven predictors are the summated construct "
           "scores defined in Table 3.1, β₀ is the intercept, β₁ to β₇ are the coefficients to be estimated, and "
           "ε is the error term. Each coefficient is the expected change in selection priority associated with a "
           "one-point increase in the corresponding construct, holding the remaining six constant. Hypotheses H1 "
           "to H7 are tested as the null that the corresponding coefficient equals zero against the one-sided "
           "alternative that it is positive, evaluated at the five per cent level.")
    w.para("Three quantities govern the interpretation. The unstandardised coefficients express effects in the "
           "units of the five-point scale and are used for substantive statements about magnitude. The "
           "standardised coefficients place the seven predictors on a common metric and are used to rank them, "
           "which is the quantity Objective 2 requires and the basis for the prioritisation recommended in "
           "Section 4.8. The adjusted coefficient of determination reports the proportion of variance in "
           "selection priority explained by the model after penalising for the number of predictors, and the "
           "F-statistic tests the joint significance of the specification.")
    w.para("Four assumptions underlying OLS are tested and the results reported in Section 4.5.1. Linearity is "
           "examined through partial regression plots of the dependent construct against each predictor. "
           "Normality of residuals is assessed through a normal probability plot and a histogram of standardised "
           "residuals, noting that at n = 800 the central limit theorem renders the inference procedures robust "
           "to moderate departures. Homoscedasticity is examined through a scatterplot of standardised residuals "
           "against standardised predicted values and confirmed formally by the Breusch–Pagan test; where "
           "heteroscedasticity is detected, heteroscedasticity-consistent standard errors are reported alongside "
           "the conventional ones. Independence of errors is checked through the Durbin–Watson statistic, with "
           "values between 1.5 and 2.5 considered acceptable. Multicollinearity has already been addressed in "
           "Section 3.4.5.")
    w.para("Section 4.5.3 reports a robustness check in which the specification is re-estimated with dummy "
           "variables for ownership type, revenue size and multi-banking status added to equation (3.1). The "
           "purpose is to establish that β₁ to β₇ are not artefacts of sample composition: if the coefficients "
           "retain their sign, significance and approximate magnitude once these controls are included, the "
           "association between the seven service constructs and selection priority is not being driven by the "
           "types of enterprise that happen to be over-represented in the sample.")

    # ============================ 3.4.7
    w.at("3.4.7. Sub-Group Difference Testing Methods (ANOVA & t-test)")
    w.para("Objective 3 asks whether selection priority differs across firm characteristics. This is tested "
           "through the five grouping variables collected in Part I of the questionnaire, summarised in "
           "Table 3.3. Each grouping variable is constructed so that its categories are mutually exclusive and "
           "collectively exhaustive, which is a precondition for analysis of variance. In particular, the "
           "guarantee product question asks for the single product the enterprise uses most frequently rather "
           "than permitting multiple selections, and ownership type separates state-owned from foreign-invested "
           "enterprises, which behave differently in this market and cannot meaningfully be pooled.")
    w.caption("Table 3.3: Grouping variables and corresponding difference tests")
    w.table([
        ["Grouping variable", "Categories", "Test applied"],
        ["Ownership type", "Private/LLC; Joint-stock (non-state); State-owned; FDI; Other",
         "One-way ANOVA with Tukey HSD post-hoc"],
        ["Annual revenue", "< 20bn; 20–100bn; 100–500bn; ≥ 500bn VND",
         "One-way ANOVA with Tukey HSD post-hoc"],
        ["Operating tenure", "< 3; 3–5; 5–10; ≥ 10 years",
         "One-way ANOVA with Tukey HSD post-hoc"],
        ["Primary guarantee product", "Tender (TG); Performance (PG); Advance payment (APG); Payment (BG); Other",
         "One-way ANOVA with Tukey HSD post-hoc"],
        ["Multi-banking status", "VietinBank only; two or more banks",
         "Independent-samples t-test"],
    ], [1.5, 2.85, 1.85], font=9.5)
    w.source("Source: compiled by the author.")
    w.para("For the four multi-category variables, one-way analysis of variance tests the null hypothesis that "
           "mean selection priority is equal across all categories. Levene's test is applied first to verify "
           "homogeneity of variance; where it is rejected, Welch's adjusted F-statistic is reported in place of "
           "the conventional one. Where the overall test is significant, Tukey's honestly significant difference "
           "procedure identifies which specific pairs of categories differ, since a significant F-statistic "
           "establishes only that some difference exists. Categories containing fewer than thirty observations "
           "are merged with a conceptually adjacent category before testing, and any such merger is reported.")
    w.para("For multi-banking status, which is binary, an independent-samples t-test compares enterprises using "
           "VietinBank exclusively with those maintaining guarantee relationships at two or more banks. The "
           "four-level response collected in the questionnaire is collapsed into this binary form for the test, "
           "while the full distribution is retained for description in Section 4.1.5. This comparison carries "
           "particular interpretive weight: enterprises that hold relationships at several banks are in a "
           "position to make the competitive comparison the dependent construct asks about, so a significant "
           "difference between the two groups bears directly on the validity of the construct as well as on the "
           "substantive question.")
    w.para("Effect sizes are reported alongside significance tests throughout. Eta squared is reported for "
           "analysis of variance and Cohen's d for the t-test. At a sample size of 800 a difference of little "
           "practical consequence can reach statistical significance, and effect sizes are what permit a "
           "difference that matters for segment strategy to be distinguished from one that is merely detectable. "
           "Hypothesis H8 is supported where significant differences with non-negligible effect sizes are found "
           "for at least one grouping variable.")


def main():
    doc = docx.Document(FILE)
    w = Writer(doc)
    write(w)
    doc.save(FILE)
    words = sum(len(p.text.split()) for p in doc.paragraphs)
    print("Da viet Chuong 3 vao", FILE, "| tong tu toan tai lieu:", words)


if __name__ == "__main__":
    main()
