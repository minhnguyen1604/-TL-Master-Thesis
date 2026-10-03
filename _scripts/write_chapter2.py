# -*- coding: utf-8 -*-
"""Viet noi dung Chuong 2 vao DTL_Master_Thesis_Draft.docx"""
import docx
from write_chapter3 import Writer

FILE = "DTL_Master_Thesis_Draft.docx"


def part_21(w):
    # ==================== 2.1
    w.at("2.1. Overview of Bank Guarantee Services in Commercial Banking")
    w.para("This section establishes what a bank guarantee is, why enterprises and banks use it, and the legal "
           "framework within which it operates in Vietnam. These foundations are necessary before the selection "
           "behaviour that forms the subject of this thesis can be theorised, because the attributes on which "
           "enterprises evaluate an issuing bank follow directly from the economic and legal structure of the "
           "instrument itself.")

    w.at("2.1.1. Nature and Economic Functions of Bank Guarantees")
    w.para("A bank guarantee is an irrevocable undertaking by a credit institution to discharge a financial "
           "obligation on behalf of its client should that client fail to perform. Circular No. 61/2024/TT-NHNN "
           "classifies the instrument as a form of credit extension, placing it alongside lending rather than "
           "among fee-based services, while Article 335 of the Civil Code 2015 provides the general civil law "
           "basis of suretyship from which the banking variant derives. Two features distinguish a bank "
           "guarantee from ordinary suretyship: the guarantor must be a licensed credit institution, and the "
           "undertaking is documentary, meaning that payment follows from the presentation of conforming "
           "documents rather than from proof that the underlying contract was in fact breached.")
    w.para("The relationship is tripartite. The principal is the enterprise whose obligation is secured; the "
           "beneficiary is the counterparty in whose favour the undertaking is issued, typically a project "
           "owner, procuring entity or supplier; and the guarantor is the bank. A defining characteristic, "
           "codified in the ICC Uniform Rules for Demand Guarantees (URDG 758) and analysed at length by "
           "Bertrams (2013), is the principle of independence: the guarantee constitutes an obligation separate "
           "from the underlying commercial contract, so the bank cannot refuse payment by invoking disputes "
           "between principal and beneficiary. This autonomy is what gives the instrument its commercial value, "
           "since the beneficiary obtains a claim against a bank rather than against a trading partner whose "
           "solvency it cannot assess.")
    w.para("From the enterprise's perspective the instrument performs three economic functions. It substitutes "
           "institutional creditworthiness for the enterprise's own, allowing a firm without an established "
           "reputation to contract with counterparties who would otherwise decline the risk. It conserves "
           "working capital, since a guarantee replaces the cash deposit that a beneficiary would otherwise "
           "demand; for a contractor simultaneously executing several projects, the liquidity released is "
           "substantial. And it transfers non-performance risk to a party equipped to price and absorb it.")
    w.para("From the bank's perspective the economics are distinctive in ways that shape competitive behaviour "
           "in this segment. At issuance the bank disburses no funds; it assumes a contingent liability recorded "
           "off the balance sheet, which converts into an on-balance-sheet claim only if the guarantee is called "
           "and the bank pays. The instrument therefore generates fee income without funding cost, producing "
           "non-interest revenue that is insulated from the interest margin compression to which lending is "
           "exposed. Because the exposure is contingent rather than drawn, it also consumes less regulatory "
           "capital than a funded loan of equivalent face value, the difference being governed by the applicable "
           "credit conversion factor. And the guarantee limit anchors the enterprise into a continuing credit "
           "relationship, creating opportunities to cross-sell deposit, payment, foreign exchange and trade "
           "finance products.")
    w.para("These properties explain why the segment is contested. A product that earns fees, consumes little "
           "capital and secures the client relationship is attractive to every commercial bank, while the "
           "enterprise faces low switching costs: obtaining a guarantee from a different bank requires no "
           "unwinding of an existing facility and can be arranged for a single transaction. Competition is "
           "therefore exercised simultaneously on fee levels, approval speed, collateral requirements and, "
           "increasingly, digital channel capability. Identifying which of these dimensions actually drives the "
           "enterprise's choice is the empirical problem this thesis addresses.")
    w.para("Carletti, Leonello and Marquez (2023) supply a necessary counterweight to this account by showing "
           "that guarantee arrangements alter the incentives of the issuing bank itself, potentially weakening "
           "the rigour of its underwriting. Their analysis is a reminder that the competitive pressures "
           "described above operate on a product whose risk is contingent and therefore easy to underprice, "
           "which is why the fee and collateral dimensions examined in this thesis carry prudential as well as "
           "commercial significance.")

    w.at("2.1.2. Main Types of Corporate Bank Guarantees")
    w.para("Vietnamese commercial banks issue guarantees across a range of product lines, each attaching to a "
           "distinct stage of the commercial or construction cycle. The distinction matters for this study "
           "because the attributes an enterprise values may depend on which product it predominantly uses, a "
           "proposition tested through the sub-group analysis described in Section 3.4.7.")
    w.caption("Table 2.1: Principal corporate bank guarantee products")
    w.table([
        ["Product", "Function", "Typical stage"],
        ["Tender guarantee (TG)",
         "Secures the bidder's commitment to maintain its bid and to sign the contract if awarded",
         "Bid submission"],
        ["Performance guarantee (PG)",
         "Secures the contractor's performance of its obligations under the signed contract",
         "Contract execution"],
        ["Advance payment guarantee (APG)",
         "Secures repayment of an advance released by the employer before work is performed",
         "Contract commencement"],
        ["Payment guarantee (BG)",
         "Secures the buyer's payment obligation towards the seller or supplier",
         "Trade settlement"],
        ["Warranty guarantee",
         "Secures rectification obligations during the defects liability period",
         "Post-completion"],
        ["Counter-guarantee / confirmed guarantee",
         "Enables a local bank to instruct a foreign bank to issue a guarantee in favour of an overseas beneficiary",
         "Cross-border trade"],
    ], [1.55, 3.35, 1.3], font=9.5)
    w.source("Source: compiled by the author from Circular 61/2024/TT-NHNN, URDG 758 and Bertrams (2013).")
    w.para("Tender and performance guarantees dominate by volume in the Vietnamese market because the Law on "
           "Bidding makes them compulsory for public procurement, a point developed in the following section. "
           "Advance payment guarantees are concentrated in construction, where employers release funds before "
           "work is performed. Payment guarantees serve trade transactions where the seller requires assurance "
           "of settlement. Counter-guarantees arise where the beneficiary is located abroad and will accept only "
           "an undertaking from a bank in its own jurisdiction, requiring the Vietnamese bank to act through a "
           "correspondent.")
    w.para("The products differ in the demands they place on the issuing bank. A tender guarantee is short-dated, "
           "modest in amount and acutely time-sensitive, since a bidder who cannot obtain the instrument before "
           "the submission deadline is excluded from the tender altogether; speed of issuance is decisive. A "
           "performance or advance payment guarantee is larger, longer in tenor and more exposed to the risk of "
           "being called, so appraisal depth and collateral terms weigh more heavily. Cross-border "
           "counter-guarantees require correspondent banking capability and familiarity with international "
           "practice. These differences provide the substantive rationale for expecting selection criteria to "
           "vary by product line, which Hypothesis H8 tests.")

    w.at("2.1.3. Legal and Regulatory Framework")
    w.para("Guarantee operations in Vietnam are governed by three overlapping bodies of rules: the general civil "
           "law of suretyship, the banking regulations issued by the State Bank of Vietnam, and the sectoral "
           "legislation that makes guarantees mandatory in particular transactions. International practice "
           "applies additionally to cross-border instruments.")
    w.para("At the general level, Article 335 of the Civil Code 2015 defines suretyship as an undertaking by a "
           "third party to perform an obligation on behalf of the obligor. At the banking level, Circular No. "
           "61/2024/TT-NHNN, effective from 1 April 2025 and replacing Circular No. 11/2022/TT-NHNN, sets out "
           "the conditions under which credit institutions may issue guarantees, the mandatory content of a "
           "letter of guarantee, the obligations of the parties and the procedures governing a demand for "
           "payment. Its most consequential innovation for the present study is the recognition of electronic "
           "guarantees. By establishing the legal basis on which an instrument issued and authenticated "
           "electronically carries the same effect as a paper original, the Circular converts digital issuance "
           "from an internal efficiency measure into a legally effective channel that a beneficiary must accept. "
           "This regulatory change is what makes digital capability a plausible determinant of bank selection "
           "rather than a matter of back-office convenience, and it is the basis for including the construct "
           "DIGITAL_CONV in the model.")
    w.para("At the sectoral level, the Law on Bidding No. 22/2023/QH15, as amended by Law No. 57/2024/QH15, "
           "requires bid security and contract performance security in public procurement, thereby creating "
           "compulsory demand for tender and performance guarantees. Decree No. 37/2015/ND-CP on construction "
           "contracts, as amended by Decree No. 35/2023/ND-CP, requires an advance payment guarantee where the "
           "employer releases an advance above the prescribed threshold. The practical consequence is that for a "
           "construction or infrastructure contractor the question is not whether to obtain guarantees but from "
           "which bank, which is precisely the decision this thesis models.")
    w.para("For cross-border transactions the ICC Uniform Rules for Demand Guarantees (URDG 758) apply where the "
           "parties incorporate them. URDG 758 codifies the independence and documentary character of the demand "
           "guarantee and regulates presentation, examination and extend-or-pay demands. Bertrams (2013) "
           "provides the authoritative treatment of these rules and of the disputes that arise under them, "
           "particularly around unfair or abusive calling, a risk that makes the issuing bank's advisory "
           "competence commercially valuable to the principal. In the Vietnamese context Le Van Dung (2021) "
           "examines the legal nature of the payment guarantee relationship at credit institutions, "
           "characterising it as a form of credit extension whose off-balance-sheet treatment does not diminish "
           "the obligations it creates.")
    w.para("Two implications follow for the research model. First, because the legal framework imposes the "
           "obligation to provide a guarantee but leaves the choice of issuer entirely free, the selection "
           "decision is genuinely discretionary and therefore worth modelling. Second, because compliance with "
           "the Law on Bidding, Circular 61/2024 and URDG 758 requires the guarantee wording to be drafted "
           "correctly, the issuing bank's legal advisory capability is a service attribute with direct financial "
           "consequences for the enterprise; this consideration is embedded in the construct STAFF_QUAL rather "
           "than treated as a separate dimension.")


def part_22(w):
    # ==================== 2.2
    w.at("2.2. Theoretical Foundations")
    w.para("Corporate bank selection is an instance of organisational buying behaviour, and no single theory "
           "accounts for all the attributes on which enterprises evaluate an issuing bank. This study therefore "
           "draws on five bodies of theory, each explaining a distinct group of attributes. Financial "
           "intermediation theory explains why the identity of the issuing bank matters at all; credit risk "
           "pricing theory explains the fee and collateral dimensions; service quality theory explains the "
           "process and personnel dimensions; relationship banking theory explains why existing ties affect "
           "terms; and technology acceptance theory explains the role of digital channels. Each theory is "
           "mapped to one or more constructs in the research model, and no construct is included without "
           "theoretical support.")

    w.at("2.2.1. Financial Intermediation & Delegated Monitoring Theory")
    w.para("Diamond (1984) explains the existence of financial intermediaries through the concept of delegated "
           "monitoring. Individual lenders face duplicated costs in monitoring a borrower, and free-riding "
           "prevents any one of them from bearing those costs alone. A bank resolves the problem by acting as "
           "the delegated monitor for many lenders simultaneously, and diversification across a large portfolio "
           "makes the bank's own promise credible without requiring each lender to monitor the bank in turn. "
           "Ramakrishnan and Thakor (1984) develop the complementary argument that intermediaries exist to "
           "produce and certify information, with reputation serving as the mechanism that makes the certified "
           "information believable.")
    w.para("Applied to guarantees, this theory identifies precisely what the instrument accomplishes. The "
           "beneficiary of a letter of guarantee cannot assess the principal's ability to perform; assessing it "
           "would require exactly the monitoring capability the beneficiary lacks. By accepting a guarantee the "
           "beneficiary substitutes a claim against an institution whose creditworthiness is publicly observable "
           "for a claim against a counterparty whose creditworthiness is not. The bank, having appraised the "
           "principal and priced the exposure, is the delegated monitor; its reputation is what makes the "
           "substitution acceptable.")
    w.para("Two consequences follow for the research model. First, the standing of the issuing bank is not an "
           "incidental attribute of the service but constitutive of the product: a guarantee is valuable to the "
           "principal only to the extent that the beneficiary accepts the issuer. Where a beneficiary declines "
           "an instrument from a lesser-known bank, the enterprise must approach an institution the beneficiary "
           "recognises, whatever the fee difference. Second, reputation in this sense is conceptually distinct "
           "from operational reliability. Reliability in the service quality literature concerns whether a "
           "provider delivers dependably on its promises to the customer; institutional reputation here concerns "
           "whether a third party who is not the customer will accept the bank's undertaking at all. The "
           "construct BANK_REP therefore enters the model on its own theoretical footing rather than as a "
           "component of service quality, a distinction that the supervisor's guidance made explicit and that "
           "Section 2.2.3 revisits.")

    w.at("2.2.2. Credit Risk Pricing & Contingent Claim Theory")
    w.para("Merton (1974) models risky corporate debt as a contingent claim whose value depends on the "
           "probability that the borrowing firm's asset value falls below its obligations, establishing the "
           "principle that the premium a lender requires is a function of default probability, exposure and the "
           "value of any security. A bank guarantee is a contingent claim in the literal sense: the bank's "
           "obligation crystallises only in the state of the world in which the principal fails to perform. The "
           "fee is accordingly the price of a contingent exposure, and cash margin and collateral reduce the "
           "loss the bank incurs in that state.")
    w.para("Stiglitz and Weiss (1981) add the information-theoretic dimension. Where the lender cannot observe "
           "borrower quality, raising the price does not simply compensate for risk but changes the composition "
           "of applicants, since the safest borrowers withdraw first; lenders therefore ration credit through "
           "non-price terms rather than clearing the market on price alone. In guarantee operations this appears "
           "as the margin ratio and collateral requirement: a bank confronting an applicant it cannot fully "
           "assess tightens security rather than simply charging more.")
    w.para("The theory supports two constructs and fixes their interpretation. Fee competitiveness (COST_COMP) "
           "is the price dimension, and collateral and margin flexibility (COLL_POLICY) is the non-price "
           "dimension; both are terms on which banks compete for the same exposure. The theory also clarifies "
           "why the two are substitutes from the enterprise's standpoint, since a low headline fee combined with "
           "a demanding margin requirement may be less attractive than the reverse once the opportunity cost of "
           "locked-up cash is taken into account. This is why the questionnaire measures the perceived "
           "reasonableness and competitiveness of the fee rather than its absolute level: the enterprise's "
           "evaluation already embeds the trade-off between the two dimensions, which a nominal fee figure would "
           "not capture.")

    w.at("2.2.3. Service Quality Theory & SERVQUAL Model")
    w.para("Parasuraman, Zeithaml and Berry (1988) developed SERVQUAL to measure perceived service quality "
           "across five dimensions: reliability, responsiveness, assurance, empathy and tangibles. The model has "
           "become the dominant framework for measuring service quality in banking, and its central premise — "
           "that quality is the customer's perception rather than a property of the service objectively defined "
           "— underlies the use of perceptual Likert measures throughout this study.")
    w.para("Two SERVQUAL dimensions translate directly into the research model. Responsiveness, the willingness "
           "to provide prompt service, corresponds to processing speed (PROC_SPEED), which in the guarantee "
           "context has unusually concrete consequences: a tender guarantee delivered after the bid deadline has "
           "no value whatever its price. Assurance, the competence of staff and their ability to inspire "
           "confidence, corresponds to staff professionalism (STAFF_QUAL), which here encompasses both technical "
           "command of guarantee mechanics and the capacity to advise on wording consistent with the Law on "
           "Bidding, Circular 61/2024 and URDG 758.")
    w.para("Zeithaml, Berry and Parasuraman (1996) extend the framework to its behavioural consequences, showing "
           "that perceived service quality translates into intentions to remain with a provider, to increase the "
           "volume of business placed with it, and to recommend it to others. This extension supplies the "
           "measurement basis for the dependent construct of this thesis, as Section 2.5 explains.")
    w.para("Two boundaries of the framework should be stated. First, SERVQUAL measures perceived quality and "
           "does not encompass price; the value construct developed in the broader literature treats price as a "
           "separate sacrifice dimension weighed against perceived benefit. Fee competitiveness in this model "
           "therefore rests on credit risk pricing theory rather than on SERVQUAL. Second, institutional "
           "reputation is not the SERVQUAL reliability dimension. Reliability concerns dependable delivery to "
           "the customer; reputation concerns acceptance of the bank's undertaking by a third-party beneficiary. "
           "Conflating them would place a construct grounded in intermediation theory inside a service quality "
           "dimension it does not belong to, and would obscure the distinct managerial levers each implies.")

    w.at("2.2.4. Relationship Banking Theory")
    w.para("Boot (2000) defines relationship banking as the provision of financial services by an intermediary "
           "that invests in obtaining customer-specific information, often proprietary, and evaluates the "
           "return on that investment across multiple products and over time. Berger and Udell (1995) provide "
           "the empirical counterpart, showing for small firm credit lines that borrowers with longer "
           "relationships obtain more favourable terms, consistent with the proposition that accumulated private "
           "information reduces the lender's uncertainty and therefore the premium it must charge.")
    w.para("The theory applies with particular force to guarantees. Appraising a guarantee application requires "
           "the bank to form a judgement about the principal's ability to perform a commercial contract, which "
           "is precisely the kind of assessment that proprietary information gathered through a continuing "
           "relationship improves. An enterprise whose payment flows, payroll and prior credit conduct are "
           "visible to the bank can be assessed more cheaply and more accurately than one presenting itself for "
           "the first time, and the saving is passed through in the form of standing limits, faster approval and "
           "more accommodating security terms. The construct RELATIONSHIP captures this mechanism, measuring "
           "relationship length, the breadth of services used and the flexibility of the guarantee limit granted.")
    w.para("The theory also predicts an asymmetry relevant to the interpretation of results. Relationship "
           "benefits accrue to the incumbent bank, creating a form of switching cost that is informational "
           "rather than contractual: a competitor offering a lower fee must also incur the appraisal cost the "
           "incumbent has already sunk. This is one reason why price competition in the segment does not "
           "eliminate incumbency advantage, and it forms part of the explanation for any positive coefficient "
           "on RELATIONSHIP in the estimated model.")

    w.at("2.2.5. Technology Acceptance Model (TAM) & Digital Banking")
    w.para("Davis (1989) proposed the Technology Acceptance Model, in which the adoption of an information "
           "technology is determined by two perceptions: usefulness, the degree to which the technology improves "
           "performance, and ease of use, the degree to which using it is free of effort. Venkatesh et al. "
           "(2003) consolidate the subsequent literature into a unified model adding social influence and "
           "facilitating conditions, while retaining performance and effort expectancy as the primary "
           "determinants.")
    w.para("Applying an individual-level adoption model in an organisational setting requires care, since the "
           "enterprise rather than the officer is the decision unit. The adaptation adopted here is that the "
           "perceptions measured are those of the officer responsible for guarantee transactions, who "
           "experiences the channel directly and whose assessment informs the enterprise's choice of issuing "
           "bank. Perceived usefulness corresponds to the reduction in elapsed time and administrative effort "
           "achieved by submitting applications online and receiving an electronically signed instrument; "
           "perceived ease of use corresponds to the accessibility of the platform and the simplicity of "
           "verification and status tracking.")
    w.para("The construct DIGITAL_CONV rests on this foundation, but its inclusion depends on a regulatory fact "
           "as much as on theory. Before Circular 61/2024/TT-NHNN, an electronically issued instrument lacked a "
           "settled legal basis and a beneficiary could reasonably insist on a paper original, so digital "
           "capability affected internal processing rather than the enterprise's choice of bank. Once electronic "
           "guarantees acquired legal effect, the channel became capable of delivering an outcome the paper "
           "process cannot match: issuance outside business hours, immediate transmission to a beneficiary in "
           "another province, and online authentication of the instrument's validity. For a contractor facing a "
           "bid deadline this is a decisive rather than a marginal difference, which is the substantive reason "
           "for retaining the construct in a deliberately parsimonious model.")
    w.para("Empirical support in the Vietnamese setting is provided by Nguyen et al. (2024), who examine service "
           "innovation in Vietnamese banks and report positive associations with customer satisfaction and "
           "loyalty. Their evidence is drawn from retail banking rather than corporate guarantee services, a "
           "limitation acknowledged in Section 2.4, but it establishes that digital service attributes influence "
           "customer outcomes in this market.")


def part_23_24(w):
    # ==================== 2.3
    w.at("2.3. Empirical Literature on Corporate Bank Selection")
    w.para("The empirical literature relevant to this thesis falls into two streams. The first examines how "
           "firms choose the bank they transact with, developed largely outside Vietnam and largely in relation "
           "to lending and general banking services. The second examines service quality and its consequences in "
           "Vietnamese banking, developed largely in retail settings. Neither stream addresses guarantee "
           "services directly, and the space between them is where this study is positioned.")

    w.at("2.3.1. International Empirical Studies")
    w.para("Turnbull and Gibbs (1989) conducted one of the earliest systematic studies of bank selection among "
           "corporate customers, examining how firms in South Africa chose and evaluated their banks. Their "
           "central finding, which has proved durable, is that corporate selection is not driven principally by "
           "price. Quality of service, the competence of the bank officers dealing with the account and the "
           "strength of the working relationship weighed at least as heavily as pricing, and the relative "
           "importance of these criteria varied with the size of the firm. The study established the corporate "
           "segment as analytically distinct from the retail segment, in which convenience and location "
           "dominate.")
    w.para("Narteh (2013) examines bank selection and patronage behaviour among small and medium enterprises in "
           "the Ghanaian banking industry, finding that selection criteria are multi-dimensional and that "
           "service delivery attributes together with bank reputation carry substantial weight in the patronage "
           "decision. Two features of this study are directly relevant here. It treats patronage as a continuous "
           "behavioural construct rather than a binary act of choice, which is the approach this thesis adopts "
           "for its dependent variable. And it confirms that the corporate findings of the earlier literature "
           "extend to the SME segment of a developing banking market, a setting closer to Vietnam than the "
           "advanced economies in which much of the literature originates.")
    w.para("Kaur et al. (2021) apply a fuzzy analytic hierarchy process to determine the bank selection criteria "
           "of SMEs, an approach that yields an explicit ranking of criteria by importance rather than a set of "
           "significance tests. Their contribution to the present study is methodological in orientation: it "
           "demonstrates that the practically useful output of this literature is a prioritised ordering of "
           "criteria, not merely a list of those that are statistically significant. This is the reasoning "
           "behind the emphasis placed on standardised coefficients in Section 3.4.6.")
    w.para("Zelie (2023) investigates the factors determining bank selection by micro and small enterprises in "
           "Ethiopia, reinforcing the finding that service quality, institutional reputation and relationship "
           "considerations operate alongside price in developing market contexts. Al-Sabbagh and Al-Khathlan "
           "(2018) move closest to the present subject by examining the factors influencing corporate clients' "
           "choice of commercial banks specifically for trade finance services. Trade finance, like guarantee "
           "issuance, is a documentary, off-balance-sheet product in which processing speed, correspondent "
           "network capability and staff expertise matter more than in ordinary lending. Their study is the "
           "nearest antecedent to this thesis, and the fact that it stops short of the guarantee product itself "
           "indicates how narrow the existing coverage is.")
    w.para("Taken together, four conclusions can be drawn from this stream. Corporate bank selection is "
           "multi-dimensional and is not reducible to price. Relationship history and institutional reputation "
           "are consistently among the stronger determinants. The relative weighting of criteria varies "
           "systematically with firm characteristics, which motivates the sub-group analysis in this thesis. And "
           "the literature has examined lending, general banking and, in one instance, trade finance, but has "
           "not isolated bank guarantee services as the object of choice.")

    w.at("2.3.2. Empirical Studies in the Vietnamese Banking Context")
    w.para("Vietnamese empirical research on banking has concentrated on service quality and its behavioural "
           "consequences, predominantly in retail settings. Phan Thi Hang Nga et al. (2024) examine the "
           "relationships among service quality, customer satisfaction and loyalty for Vietnamese SMEs, "
           "confirming that the service quality to loyalty pathway operates in this market and that satisfaction "
           "mediates it. Ho Dinh Phi et al. (2023) examine the same pathway with corporate reputation as an "
           "additional mediating construct, a finding that supports the decision in this thesis to model "
           "reputation as a determinant in its own right rather than folding it into service quality. Nguyen et "
           "al. (2024) examine service innovation in Vietnamese retail banks and report positive associations "
           "with satisfaction and loyalty, providing the local evidence base for the digital construct.")
    w.para("Research addressing bank guarantees specifically in Vietnam has been predominantly legal and "
           "qualitative in character. Le Van Dung (2021) analyses the nature of the payment guarantee "
           "relationship at credit institutions, clarifying its status as credit extension and examining the "
           "obligations arising between the parties. Work of this kind is indispensable for understanding the "
           "instrument, and Section 2.1 draws on it, but it does not measure the behaviour of enterprises "
           "choosing between issuing banks.")
    w.para("The Vietnamese stream therefore exhibits a consistent division. Studies employing quantitative "
           "methods examine retail banking or general service quality; studies examining guarantees employ legal "
           "and qualitative methods. No study located during this review applies a quantitative selection model "
           "to the corporate guarantee product in Vietnam.")

    # ==================== 2.4
    w.at("2.4. Research Gaps")
    w.para("Four gaps emerge from the review, each of which the present study addresses directly.")
    w.caption("Table 2.2: Research gaps and the response of this study")
    w.table([
        ["Gap", "Evidence from the literature", "Response of this study"],
        ["1. Guarantee services are not isolated as the object of selection",
         "Corporate selection studies examine lending and general banking (Turnbull and Gibbs, 1989; Narteh, "
         "2013; Kaur et al., 2021; Zelie, 2023); the closest, Al-Sabbagh and Al-Khathlan (2018), reaches trade "
         "finance but not guarantees",
         "Models selection of the guarantee product specifically, with constructs defined against its economic "
         "and legal characteristics"],
        ["2. No integration of price competitiveness with digital channel capability",
         "Digital attributes appear in Vietnamese retail studies (Nguyen et al., 2024) but not alongside fee and "
         "collateral terms in a single corporate selection model",
         "Estimates fee competitiveness and digital eFAST convenience within one equation, under the new legal "
         "regime for e-guarantees"],
        ["3. Absence of sub-group comparison in the Vietnamese context",
         "The international literature indicates criteria vary by firm characteristics, but Vietnamese studies "
         "report pooled results",
         "Tests differences across ownership type, revenue size, operating tenure, primary product line and "
         "multi-banking status"],
        ["4. Methodological division between quantitative and guarantee-focused work in Vietnam",
         "Vietnamese quantitative studies address retail service quality; Vietnamese guarantee studies are legal "
         "and qualitative (Le Van Dung, 2021)",
         "Applies a quantitative multivariate model to the corporate guarantee segment at a single large "
         "commercial bank"],
    ], [1.45, 2.65, 2.1], font=9)
    w.source("Source: compiled by the author.")
    w.para("A further consideration reinforces the timeliness of the study. The literature on which the model "
           "draws predates Circular No. 61/2024/TT-NHNN, which came into effect on 1 April 2025. Any empirical "
           "account of guarantee selection formed before that date could not incorporate the legal effectiveness "
           "of electronic issuance, because electronic issuance did not yet possess it. The regulatory change "
           "creates a genuine opportunity to test whether digital capability has become a determinant of bank "
           "choice in this segment rather than a matter of internal efficiency, and that test is not available "
           "in the existing literature.")


def part_25(w):
    # ==================== 2.5
    w.at("2.5.1. Conceptual Research Framework")
    w.para("The conceptual framework brings together the five theoretical strands of Section 2.2 and the "
           "empirical patterns of Section 2.3 into a single-stage model. Seven independent constructs are "
           "hypothesised to be associated with one dependent construct, selection priority. The model contains "
           "no mediating or moderating variables: the relationships are specified as direct, which is why "
           "multiple regression rather than structural equation modelling is the appropriate estimation "
           "technique. Differences across firm characteristics are handled through sub-group comparison rather "
           "than through interaction terms in the main equation, preserving the interpretability of the "
           "coefficients on which the managerial conclusions depend.")
    w.para("Figure 2.1 presents the framework. The seven constructs are arrayed on the left, each with its "
           "theoretical source; selection priority is on the right; and the five grouping variables appear "
           "beneath as the basis of the difference tests.")
    w.caption("Figure 2.1: Conceptual research framework")
    w.table([
        ["Independent construct", "Theoretical source", "Expected sign", "Dependent construct"],
        ["COST_COMP – Price Competitiveness", "Merton (1974); Stiglitz and Weiss (1981)", "+",
         "DEC\n\nSelection Priority &\nPatronage Intention\n\n(Zeithaml, Berry and\nParasuraman, 1996;\nOliver, 1999)"],
        ["PROC_SPEED – Processing Speed", "Parasuraman, Zeithaml and Berry (1988)", "+", ""],
        ["DIGITAL_CONV – Digital eFAST Convenience", "Davis (1989); Venkatesh et al. (2003)", "+", ""],
        ["BANK_REP – Bank Reputation", "Diamond (1984); Ramakrishnan and Thakor (1984)", "+", ""],
        ["RELATIONSHIP – Relationship & Limits", "Boot (2000); Berger and Udell (1995)", "+", ""],
        ["STAFF_QUAL – Staff Professionalism", "Parasuraman, Zeithaml and Berry (1988)", "+", ""],
        ["COLL_POLICY – Collateral & Margin Flexibility", "Merton (1974); Stiglitz and Weiss (1981)", "+", ""],
        ["Grouping variables: ownership type · revenue size · operating tenure · primary guarantee product · "
         "multi-banking status", "", "H8", ""],
    ], [2.0, 2.05, 0.65, 1.55], font=8.5)
    w.source("Source: constructed by the author from the theoretical framework in Section 2.2.")

    w.at("2.5.2. Hypothesis Development")
    w.para("Eight hypotheses follow from the framework. Seven concern the association between an individual "
           "construct and selection priority; the eighth concerns differences across firm characteristics. All "
           "seven directional hypotheses are positive, which requires a word of explanation. In a model of this "
           "kind a negative sign would be expected for any construct measuring a sacrifice the customer incurs. "
           "No construct here is measured in that form: the fee construct records the enterprise's evaluation of "
           "how reasonable and competitive the pricing is, not the level of the fee itself, so a higher score "
           "denotes a more favourable rather than a more onerous condition. Uniform positive signs are therefore "
           "a property of the measurement direction, not an assumption that price is irrelevant.")
    for h, text in [
        ("H1", "Price Competitiveness (COST_COMP) is positively associated with corporate selection priority. "
               "Credit risk pricing theory treats the guarantee fee as the price of a contingent exposure "
               "(Merton, 1974), and enterprises facing low switching costs can compare that price across banks. "
               "Where the fee schedule is judged reasonable relative to the service received and competitive "
               "against alternatives, the enterprise has a direct financial reason to concentrate its guarantee "
               "business at that bank. The international evidence that price is not the dominant criterion "
               "(Turnbull and Gibbs, 1989) does not imply that it is immaterial; it implies that its "
               "coefficient should be positive but need not be the largest."),
        ("H2", "Processing Speed (PROC_SPEED) is positively associated with corporate selection priority. The "
               "responsiveness dimension of SERVQUAL (Parasuraman, Zeithaml and Berry, 1988) acquires unusual "
               "force in this product because guarantee deadlines are externally imposed. A tender guarantee "
               "that arrives after the submission deadline has no value at any price, and an advance payment "
               "guarantee delayed beyond the contract milestone postpones the release of funds to the "
               "contractor. Speed here is not a matter of comfort but of whether the enterprise can transact at "
               "all."),
        ("H3", "Digital eFAST Convenience (DIGITAL_CONV) is positively associated with corporate selection "
               "priority. The Technology Acceptance Model (Davis, 1989; Venkatesh et al., 2003) predicts that "
               "perceived usefulness and ease of use drive adoption, and Circular No. 61/2024/TT-NHNN has made "
               "the electronic channel legally effective rather than merely internally convenient. Online "
               "submission outside business hours, digitally signed issuance and online authentication "
               "collapse the elapsed time of the transaction, which under H2 is precisely what the enterprise "
               "values."),
        ("H4", "Bank Reputation (BANK_REP) is positively associated with corporate selection priority. "
               "Delegated monitoring theory (Diamond, 1984; Ramakrishnan and Thakor, 1984) establishes that the "
               "instrument works only because the beneficiary accepts the issuer's undertaking in place of the "
               "principal's. Reputation is therefore not an ancillary attribute but a condition of the "
               "product's usefulness, and Ho Dinh Phi et al. (2023) provide Vietnamese evidence that "
               "institutional reputation operates on customer outcomes in this market."),
        ("H5", "Relationship Banking and Limits (RELATIONSHIP) is positively associated with corporate "
               "selection priority. Relationship banking theory (Boot, 2000) holds that proprietary information "
               "accumulated through a continuing relationship lowers the cost of assessment, and Berger and "
               "Udell (1995) show empirically that longer relationships secure more favourable credit terms. "
               "Applied to guarantees, an enterprise whose accounts and credit conduct are already visible to "
               "the bank obtains standing limits and faster approval, both of which raise the priority it "
               "assigns to that bank."),
        ("H6", "Staff Professionalism (STAFF_QUAL) is positively associated with corporate selection priority. "
               "The assurance dimension of SERVQUAL concerns staff competence and the confidence it inspires. "
               "In guarantee operations this has a specific financial dimension: an officer who can draft "
               "wording consistent with the Law on Bidding, Circular 61/2024 and URDG 758 reduces the "
               "enterprise's exposure to an unfair or abusive call on the instrument, a risk documented in the "
               "guarantee literature (Bertrams, 2013)."),
        ("H7", "Collateral and Margin Flexibility (COLL_POLICY) is positively associated with corporate "
               "selection priority. Stiglitz and Weiss (1981) show that lenders facing imperfect information "
               "ration through non-price terms rather than price alone, and cash margin is the principal such "
               "term in guarantee operations. A margin requirement immobilises working capital, so the "
               "opportunity cost of a demanding margin can exceed the fee itself; flexibility on this dimension "
               "is therefore a substantive competitive instrument, not an administrative detail."),
    ]:
        w.para(h + ". " + text, first_line=0, indent=0.25, after=8)
    w.para("H8. Corporate selection priority differs significantly across firm characteristics, specifically "
           "ownership type, revenue size, operating tenure, and primary guarantee product line. The "
           "international literature indicates that the weighting of selection criteria varies with firm "
           "characteristics (Turnbull and Gibbs, 1989; Kaur et al., 2021), and Section 2.1.2 established that "
           "the guarantee products themselves impose different demands on the issuing bank: a tender guarantee "
           "is decided on speed, while a performance or advance payment guarantee is decided on appraisal and "
           "security terms. State-owned enterprises, private firms and foreign-invested firms face different "
           "procurement regimes and different counterparties, which is why the questionnaire separates them "
           "rather than pooling state and foreign ownership into a single category.",
           first_line=0, indent=0.25, after=8)
    w.para("The eight hypotheses are tested through the procedures set out in Chapter 3: H1 to H7 through the "
           "coefficients of the regression model specified in equation (3.1), and H8 through the analysis of "
           "variance and t-test procedures described in Section 3.4.7.")


def main():
    doc = docx.Document(FILE)
    w = Writer(doc)
    part_21(w); part_22(w); part_23_24(w); part_25(w)
    doc.save(FILE)
    print("Da viet Chuong 2 vao", FILE)


if __name__ == "__main__":
    main()
