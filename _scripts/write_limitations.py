# -*- coding: utf-8 -*-
"""Viet muc 4.10 Research Limitations and Suggestions for Future Research."""
import docx
from write_chapter3 import Writer

FILE = "DTL_Master_Thesis_Draft.docx"


def main():
    doc = docx.Document(FILE)
    w = Writer(doc)
    w.at("4.10. Research Limitations and Suggestions for Future Research")
    w.para("The findings reported above should be read against the limitations of the research design. Six are "
           "set out here, each followed by the direction it suggests for further work. They are stated as "
           "boundaries on what this study can establish rather than as defects, since most follow from choices "
           "made deliberately in Chapter 3 to keep the study feasible within the scope of a master's thesis.")

    w.para("First, the sample is drawn entirely from enterprises that are already customers of VietinBank. "
           "Enterprises that considered the bank and chose a competitor, and those that formerly used it and "
           "left, are absent from the frame. The consequence is that the model explains the priority an existing "
           "client assigns to VietinBank and the share of guarantee business it concentrates there; it does not "
           "explain first-time bank choice across the market. The dependent construct and the managerial "
           "implications in Section 4.8 are framed accordingly, and the item DEC4, which places VietinBank "
           "explicitly against alternatives, mitigates but does not remove the restriction. The natural "
           "extension is a comparative design surveying enterprises that obtain guarantees from other banks, or "
           "a study of enterprises that have switched issuing bank, either of which would identify the "
           "determinants of acquisition rather than retention.", first_line=0.25)

    w.para("Second, the design is cross-sectional. All variables are measured at one point in time, so the "
           "estimated relationships are associations and temporal precedence cannot be established. A favourable "
           "evaluation of processing speed may raise selection priority, but an enterprise that has already "
           "concentrated its business at the bank may equally evaluate that bank's service more generously. "
           "Disentangling the two requires a longitudinal design in which perceptions are measured before "
           "subsequent allocation decisions are observed, or a design using instrumental variables. Both are "
           "beyond the scope of this thesis and are recommended for future work.", first_line=0.25)

    w.para("Third, sampling was non-probabilistic. Because a complete list of qualifying enterprises could not "
           "be released, stratified convenience sampling was used, with regional quotas approximating the "
           "population distribution. Selection probabilities are therefore unknown, sampling error cannot be "
           "estimated, and inference rests on the assumption that respondents do not differ systematically from "
           "non-respondents. Replication with a probability sample, which would require the cooperation of the "
           "bank in drawing from its client register under appropriate confidentiality safeguards, would permit "
           "stronger inference.", first_line=0.25)

    w.para("Fourth, all constructs are measured through the perceptions of a single respondent per enterprise, "
           "using a single instrument administered at one time. This raises the possibility of common method "
           "bias (Podsakoff et al., 2003), and administration through the bank's own relationship channel adds "
           "a risk of social desirability bias. The procedural remedies described in Section 3.3.3 were applied, "
           "and the descriptive evidence in Section 4.1 was inspected for the compressed dispersion "
           "characteristic of the problem. Future research could triangulate the perceptual measures against "
           "objective transaction records — realised guarantee volumes, fee rates actually charged, and observed "
           "channel usage — which would test whether stated priority corresponds to revealed behaviour.",
           first_line=0.25)

    w.para("Fifth, the study covers a single bank. VietinBank's scale, ownership structure, branch network and "
           "digital maturity shape the evaluations reported here, and the weights estimated for the seven "
           "constructs need not hold at banks differing on those dimensions. A smaller private bank cannot "
           "compete on the reputation dimension in the same way and may compete more aggressively on price and "
           "flexibility, which would plausibly alter the ranking. A multi-bank study, ideally spanning both "
           "state-owned and private institutions, would establish which findings are general and which are "
           "specific to a large incumbent.", first_line=0.25)

    w.para("Sixth, the model is deliberately parsimonious. Seven constructs were retained, and attributes such "
           "as branch network density, correspondent banking reach for cross-border guarantees, and the "
           "capacity to structure bespoke guarantee wording were excluded to protect degrees of freedom and to "
           "limit multicollinearity. Their exclusion is a modelling decision, not a finding that they are "
           "irrelevant, and the adjusted coefficient of determination reported in Section 4.5.1 indicates how "
           "much variation remains unexplained. Future work could examine these attributes, particularly within "
           "the sub-population of exporting and foreign-invested enterprises for whom correspondent capability "
           "is likely to matter most. A further extension would introduce customer satisfaction as a mediating "
           "construct and estimate the model through structural equation modelling, following the approach of "
           "Phan Thi Hang Nga et al. (2024) and Ho Dinh Phi et al. (2023); the single-stage specification used "
           "here was chosen because the framework contains no mediating variable, but a mediated structure is a "
           "plausible alternative worth testing.", first_line=0.25)

    w.para("A final observation concerns timing rather than design. The survey was conducted shortly after "
           "Circular No. 61/2024/TT-NHNN took effect, at a point when adoption of electronic guarantees was "
           "still developing across both banks and beneficiaries. The coefficient estimated for DIGITAL_CONV "
           "therefore reflects an early stage of that transition. Repeating the study once electronic issuance "
           "has become routine would show whether digital capability functions as a durable point of "
           "differentiation or, as competitors close the gap, settles into a baseline expectation that no "
           "longer distinguishes one bank from another.", first_line=0.25)

    w.para("[Cần bổ sung sau khi có kết quả: các hạn chế phát sinh từ chính kết quả thực nghiệm — ví dụ nhân tố "
           "bị gộp trong EFA, biến quan sát phải loại do hệ số tải thấp, nhóm có cỡ mẫu nhỏ phải gộp trước khi "
           "chạy ANOVA, hoặc phần phương sai chưa giải thích được nếu R² hiệu chỉnh thấp hơn kỳ vọng.]",
           italic=True, first_line=0)

    doc.save(FILE)
    print("Da viet muc 4.10 Research Limitations")


if __name__ == "__main__":
    main()
