# -*- coding: utf-8 -*-
"""Viet noi dung Chuong 1 vao DTL_Master_Thesis_Draft.docx"""
import docx
from write_chapter3 import Writer

FILE = "DTL_Master_Thesis_Draft.docx"


def write(w):
    # ==================== 1.1
    w.at("1.1. Research Rationales & Background")
    w.para("In modern commercial banking the bank guarantee occupies an unusual position. It is a credit "
           "instrument that ordinarily consumes no funds, a fee-earning service whose fee is in substance a risk "
           "premium, and a product whose usefulness depends not on the satisfaction of the paying customer but "
           "on its acceptance by a third party who is not a customer at all. These characteristics, examined in "
           "detail in Chapter 2, make the segment economically attractive to banks and strategically important "
           "to the enterprises that depend on it.")
    w.para("For Vietnamese enterprises, demand for guarantees is created largely by law rather than by "
           "commercial preference. The Law on Bidding No. 22/2023/QH15, as amended by Law No. 57/2024/QH15, "
           "requires bid security and contract performance security in public procurement. Decree No. "
           "37/2015/ND-CP on construction contracts, as amended by Decree No. 35/2023/ND-CP, requires an advance "
           "payment guarantee where the employer releases an advance above the prescribed threshold. In "
           "cross-border trade, guarantees issued subject to the ICC Uniform Rules for Demand Guarantees "
           "(URDG 758) substitute for cash deposits and release working capital that would otherwise be "
           "immobilised. A construction contractor or trading enterprise operating at scale in Vietnam does not "
           "decide whether to obtain guarantees; it decides only which bank will issue them.")
    w.para("That decision is genuinely open. Unlike a term loan, a guarantee does not lock the enterprise into a "
           "multi-year relationship: an instrument can be arranged for a single contract, and turning to a "
           "different bank for the next one requires no unwinding of an existing facility. Switching costs are "
           "therefore low, and banks compete for the same business simultaneously on fee levels, on the speed "
           "with which the instrument is issued, on the cash margin and collateral demanded, on the "
           "institutional standing that determines whether beneficiaries accept the letter, and increasingly on "
           "the digital channel through which the transaction is conducted. What is not known, in the Vietnamese "
           "market, is the relative weight enterprises actually place on these competing dimensions.")
    w.para("Two recent developments make the question more pressing than it would have been a few years ago. "
           "The first is regulatory. Circular No. 61/2024/TT-NHNN, effective from 1 April 2025 and replacing "
           "Circular No. 11/2022/TT-NHNN, established a complete legal framework for electronic guarantees. "
           "Before that change, a bank could digitise its internal workflow but could not issue an instrument "
           "that a beneficiary was obliged to accept without a paper original; digital capability was an "
           "efficiency measure. After it, an electronically issued and authenticated guarantee carries the same "
           "legal effect as a paper one, which converts the digital channel into something an enterprise can "
           "choose a bank for. Whether enterprises in fact do so is an empirical question that could not have "
           "been asked before 2025.")
    w.para("The second is competitive. Sustained pressure on net interest margins has pushed Vietnamese "
           "commercial banks to expand non-interest income, and guarantee fees are among the most attractive "
           "sources available: they are recurring, they require no funding, and the underlying exposure consumes "
           "less regulatory capital than funded lending of equivalent face value. The predictable consequence is "
           "intensified competition for guarantee business, expressed through fee discounting, faster approval "
           "and more accommodating security terms. Competing on all dimensions at once is expensive, and a bank "
           "that does not know which dimension actually moves the client's decision will allocate that spending "
           "inefficiently.")
    w.para("Against this background the academic literature is notably thin. As Section 2.3 documents, research "
           "on corporate bank selection has concentrated on lending and general banking services, reaching trade "
           "finance in one instance but not guarantee issuance; Vietnamese quantitative research has "
           "concentrated on retail service quality; and Vietnamese research on guarantees has been predominantly "
           "legal and qualitative. No study identified in this review applies a quantitative selection model to "
           "the corporate guarantee product in Vietnam, and none incorporates digital channel capability under "
           "the legal regime now in force.")
    w.para("This thesis addresses that gap. It examines the factors associated with corporate customers' "
           "decision to choose Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank) for bank "
           "guarantee services, estimating the direction and relative importance of seven service and "
           "relationship attributes, and testing whether their importance differs across types of enterprise and "
           "types of guarantee product.")

    # ==================== 1.2
    w.at("1.2. Research Problem & Industry Context at VietinBank")
    w.para("VietinBank is one of Vietnam's largest commercial banks, operating a nationwide corporate banking "
           "network of 155 domestic branches together with the VietinBank eFAST digital platform for corporate "
           "clients. Its scale in the corporate segment makes it a suitable setting for this study on two "
           "grounds: the population of corporate guarantee users is large enough to support the sample "
           "requirements set out in Section 3.3.2, and the bank competes across the full range of guarantee "
           "products, from short-dated tender guarantees for domestic contractors to counter-guarantees "
           "supporting cross-border trade.")
    w.para("[Cần bổ sung số liệu: quy mô dư nợ bảo lãnh và tốc độ tăng trưởng của VietinBank trong 3–5 năm gần "
           "nhất; thu nhập phí bảo lãnh và tỷ trọng trong tổng thu nhập ngoài lãi; vị thế so với các ngân hàng "
           "cạnh tranh trong phân khúc bảo lãnh khách hàng doanh nghiệp. Nguồn: Báo cáo thường niên và Báo cáo "
           "tài chính hợp nhất đã kiểm toán của VietinBank, thuyết minh về cam kết ngoại bảng; báo cáo ngành của "
           "Ngân hàng Nhà nước. Bổ sung tài liệu này vào danh mục tham khảo sau khi trích dẫn.]", italic=True)
    w.para("The managerial problem arises from the combination of these two facts. The bank competes on several "
           "dimensions at once, and each carries a real cost. Fee discounting reduces revenue on every "
           "transaction to which it applies. Compressing appraisal turnaround requires either additional "
           "underwriting capacity or a relaxation of review depth. Reducing the cash margin increases the "
           "unsecured portion of the contingent exposure. Extending the digital channel requires sustained "
           "investment in platform development, authentication infrastructure and integration with beneficiary "
           "verification. Each of these is a legitimate competitive instrument, and each consumes resources that "
           "cannot simultaneously be spent on the others.")
    w.para("Deciding among them requires evidence about what corporate clients actually weigh, and that evidence "
           "is not currently available. Internal transaction records reveal what enterprises did — which "
           "guarantees were issued, at what fee, through which channel — but not why they brought the business "
           "to VietinBank rather than to a competitor. A client who would have remained regardless of a fee "
           "discount is indistinguishable in the transaction data from one whom the discount retained. The "
           "information needed to allocate competitive spending efficiently is perceptual, concerns the "
           "comparison the enterprise makes between banks, and can be obtained only by asking.")
    w.para("A second dimension of the problem concerns segmentation. A uniform competitive posture across the "
           "entire corporate base is unlikely to be efficient if, as the international literature suggests, "
           "selection criteria vary with firm characteristics. A state-owned enterprise operating under public "
           "procurement rules, a private contractor managing tender deadlines, and a foreign-invested firm "
           "requiring counter-guarantees abroad plausibly weigh speed, price and reputation differently. If such "
           "differences exist and are of material size, they justify differentiated product and pricing "
           "strategies; if they do not, a uniform approach is the more efficient choice. The bank presently has "
           "no empirical basis on which to make that determination.")
    w.para("The research problem this thesis addresses can therefore be stated precisely: VietinBank competes "
           "for corporate guarantee business across several costly dimensions without measured evidence of which "
           "dimensions drive corporate selection priority, and without evidence of whether those drivers differ "
           "across the segments of its corporate client base.")

    # ==================== 1.3 / 1.4 lead-ins
    w.at("1.3. Research Objectives")
    w.para("The objectives below translate the research problem into measurable aims. The general objective "
           "states the overall purpose of the study; the four specific objectives decompose it into components "
           "that can each be addressed by an identifiable analytical procedure, and each is matched to one "
           "research question in Section 1.4 and to one group of recommendations in Section 4.8.")

    w.at("1.4. Research Questions")
    w.para("Each research question corresponds to the specific objective bearing the same number and is answered "
           "by a designated part of the empirical analysis: RQ1 by the factor structure and correlation results, "
           "RQ2 by the standardised regression coefficients, RQ3 by the sub-group difference tests, and RQ4 by "
           "the managerial discussion that follows from them.")

    # ==================== 1.5
    w.at("1.5. Scope and Boundaries of the Study")
    w.para("The scope of the study is defined along four dimensions, and the boundaries are stated explicitly so "
           "that the findings are not extended beyond what the design can support.")
    w.para("In subject matter, the study is confined to bank guarantee services for corporate customers. Other "
           "corporate banking products — term lending, working capital finance, documentary credits, payments "
           "and foreign exchange — enter the analysis only insofar as their prior use constitutes part of the "
           "banking relationship measured by the construct RELATIONSHIP. Guarantee services for individual and "
           "household customers are outside the scope, since the corporate product differs in appraisal "
           "process, documentation and legal framework.")
    w.para("In content, the model examines seven determinants: price competitiveness, processing speed, digital "
           "eFAST convenience, bank reputation, relationship and limits, staff professionalism, and collateral "
           "and margin flexibility. This set follows the supervisor's guidance on model parsimony and is "
           "justified construct by construct in Section 2.2. Attributes not included — among them branch network "
           "density, correspondent banking reach and bespoke structuring capability — are not thereby asserted "
           "to be irrelevant; they were excluded to keep the model within a size that the available degrees of "
           "freedom and the risk of multicollinearity can support, and Section 4.10 identifies them as "
           "candidates for future research.")
    w.para("In space, the study covers corporate customers of VietinBank across its 155 domestic branches. "
           "Findings therefore describe one large Vietnamese commercial bank; they are not a market-wide "
           "estimate, and generalisation to banks of materially different scale, ownership or digital maturity "
           "requires replication.")
    w.para("In time, the survey is conducted at a single point, and the legal framework described is that in "
           "force following the entry into effect of Circular No. 61/2024/TT-NHNN on 1 April 2025. The "
           "cross-sectional design means the study measures association rather than causation and cannot "
           "capture change in selection criteria over time.")
    w.para("One boundary requires particular emphasis because it shapes how the results must be read. The sample "
           "consists of enterprises that are already customers of VietinBank. Enterprises that considered the "
           "bank and selected a competitor, and those that formerly used it and left, are not represented. The "
           "dependent construct accordingly measures the priority an existing customer assigns to VietinBank "
           "relative to the other banks it uses or has considered, and the study's contribution lies in "
           "explaining retention and the share of guarantee business a client concentrates at the bank, rather "
           "than in explaining first-time bank choice across the market as a whole. The managerial implications "
           "in Section 4.8 are framed within that boundary, and Section 4.10 proposes a comparative design "
           "including non-customers as the appropriate extension.")

    # ==================== 1.6
    w.at("1.6. Significance & Contributions of the Study")
    w.para("The study offers contributions of three kinds.")
    w.para("Theoretically, it isolates bank guarantee services as the object of the selection decision. On the "
           "basis of the review in Section 2.3, no previous quantitative study in Vietnam has done so: the "
           "corporate selection literature has examined lending and general banking services, and the "
           "Vietnamese guarantee literature has been legal and qualitative in method. The study also "
           "demonstrates that a coherent model of this decision requires drawing on several theoretical "
           "traditions simultaneously, since service quality theory alone cannot account for the price and "
           "security dimensions, and intermediation theory alone cannot account for the process and personnel "
           "dimensions. In doing so it clarifies a distinction that is frequently blurred in applied work: "
           "institutional reputation, understood as the willingness of a third-party beneficiary to accept the "
           "issuer's undertaking, is not the reliability dimension of service quality and should not be modelled "
           "as one. Finally, the study provides the first opportunity to test whether digital channel capability "
           "influences corporate bank selection under a legal regime in which electronic guarantees are "
           "effective, a condition that did not exist before April 2025.")
    w.para("Practically, the study converts a set of competing managerial intuitions into a ranked and "
           "quantified ordering. Because the seven constructs are estimated within a single equation, the "
           "standardised coefficients indicate not merely which attributes matter but which matter more, which "
           "is the form in which the finding becomes usable for allocating a finite competitive budget. The "
           "sub-group analysis further establishes whether a uniform competitive posture is appropriate or "
           "whether differentiated strategies by ownership type, firm size and guarantee product line are "
           "warranted. Recommendations addressed to the State Bank of Vietnam follow from the same evidence, "
           "concerning the consistent implementation of the electronic guarantee framework and the transparency "
           "of fee disclosure in the corporate guarantee market.")
    w.para("Methodologically, the study develops and validates a measurement instrument specific to corporate "
           "guarantee selection. Existing scales were designed for retail banking or for general service "
           "quality and require substantial adaptation before they can measure attributes such as beneficiary "
           "acceptance of a letter of guarantee, margin flexibility or the legal competence of guarantee "
           "wording. The thirty-two-item instrument reported in Section 3.2, together with its reliability and "
           "factor analysis results, provides a validated starting point for subsequent research on this "
           "product, whether at other Vietnamese banks or in comparative work across banks.")

    # ==================== 1.7
    w.at("1.7. Structure of the Thesis")
    w.para("The thesis is organised in four chapters.")
    for t in [
        "Chapter 1 establishes the rationale for the study, sets out the research problem in the context of "
        "VietinBank's corporate guarantee operations, states the general and specific objectives and the "
        "corresponding research questions, defines the scope and boundaries, and identifies the contributions "
        "the study makes.",
        "Chapter 2 reviews the literature and develops the theoretical framework. It examines the nature, "
        "product range and legal framework of bank guarantees in Vietnam; sets out the five theories on which "
        "the model rests; reviews the international and Vietnamese empirical literature on corporate bank "
        "selection; identifies four research gaps; and derives the conceptual framework together with the eight "
        "hypotheses.",
        "Chapter 3 presents the research methodology. It describes the research design and the seven-step "
        "analytical procedure, the operationalisation of the eight constructs and their mapping to the "
        "thirty-two-item questionnaire, the target population, sampling strategy and sample size "
        "determination, and the econometric methods applied, including the regression specification and the "
        "sub-group difference tests.",
        "Chapter 4 reports and discusses the empirical results. It presents the sample profile, the reliability "
        "and factor analysis results, the correlation and multicollinearity diagnostics, the regression "
        "estimates and hypothesis tests together with a robustness check, and the sub-group comparisons. It "
        "then discusses the findings against the literature, sets out managerial recommendations for VietinBank "
        "and policy recommendations for the State Bank of Vietnam, and states the limitations of the study "
        "together with directions for future research."]:
        w.bullet(t)
    w.para("References and four appendices containing the full statistical output follow Chapter 4.")


def main():
    doc = docx.Document(FILE)
    w = Writer(doc)
    write(w)
    doc.save(FILE)
    print("Da viet Chuong 1 vao", FILE)


if __name__ == "__main__":
    main()
