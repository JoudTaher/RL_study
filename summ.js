const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, PageBreak, Tab, TabStopType, TabStopPosition,
  Header, Footer
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, color: "1F3864", font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, color: "2E5FA3", font: "Arial" })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 24, color: "1F3864", font: "Arial" })]
  });
}

function body(text, options = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 276 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...options })]
  });
}

function bodyMixed(runs) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 276 },
    children: runs.map(r => new TextRun({ size: 22, font: "Arial", ...r }))
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function spacer() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 80, after: 80 } });
}

function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E5FA3", space: 1 } },
    spacing: { before: 200, after: 200 },
    children: [new TextRun("")]
  });
}

function infoBox(label, content) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: "EBF3FB", type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [new TextRun({ text: label, bold: true, size: 20, font: "Arial", color: "1F3864" })]
              }),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                spacing: { before: 40, after: 40, line: 276 },
                children: [new TextRun({ text: content, size: 20, font: "Arial" })]
              })
            ]
          })
        ]
      })
    ]
  });
}

function makeTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: "1F3864", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, bold: true, size: 18, font: "Arial", color: "FFFFFF" })]
      })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders,
      width: { size: colWidths[ci], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? "F5F9FF" : "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: cell, size: 18, font: "Arial" })]
      })]
    }))
  }));

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E5FA3" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1F3864", space: 1 } },
            spacing: { after: 100 },
            children: [new TextRun({ text: "Literature Review — Oracle-RLAIF | Reinforcement Learning Course", size: 18, font: "Arial", color: "666666" })]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: "1F3864", space: 1 } },
            alignment: AlignmentType.RIGHT,
            spacing: { before: 100 },
            children: [
              new TextRun({ text: "Page ", size: 18, font: "Arial", color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: "666666" })
            ]
          })
        ]
      })
    },
    children: [

      // ─── TITLE BLOCK ───────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 40 },
        children: [new TextRun({ text: "Literature Review", bold: true, size: 44, font: "Arial", color: "1F3864" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Oracle-RLAIF: An Improved Fine-Tuning Framework", bold: true, size: 32, font: "Arial", color: "2E5FA3" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "for Multi-modal Video Models through Reinforcement Learning from Oracle Ranking Feedback", size: 28, font: "Arial", color: "2E5FA3" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Shi et al. (2025) — Lawrence Livermore National Laboratory", size: 20, font: "Arial", color: "888888", italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 280 },
        children: [new TextRun({ text: "Submitted for: Reinforcement Learning Course", size: 20, font: "Arial", color: "888888" })]
      }),

      divider(),

      // ─── 1. ABSTRACT ───────────────────────────────────────────────────────────
      h1("1. Abstract"),
      body("This literature review critically examines Oracle-RLAIF (Shi et al., 2025), a state-of-the-art framework for aligning large video-language models (VLMs) through reinforcement learning from AI-generated ordinal feedback. The work targets a fundamental bottleneck in modern RL-based alignment pipelines: the dependency on expensive, domain-specific reward models that must assign calibrated scalar scores to arbitrary prompt–response pairs. Oracle-RLAIF resolves this by substituting the reward model with a drop-in Oracle ranker — any model capable of producing a relative ordering of candidate responses — thereby decoupling the alignment pipeline from reward magnitude calibration. To leverage ordinal feedback within a policy gradient framework, the authors introduce GRPOrank, a rank-adapted extension of Group Relative Policy Optimization (GRPO) that replaces scalar advantage estimation with a normalized Discounted Cumulative Gain (nDCG) penalty, ensuring position-sensitive, bounded, and zero-sum learning signals."),
      spacer(),
      body("This review grounds each component of Oracle-RLAIF in its RL theoretical context, covering Markov Decision Process (MDP) formulation, policy gradient methods, advantage estimation, and the evolution from PPO through GRPO to GRPOrank. Empirical evaluation across MSVD-QA, MSRVTT-QA, ActivityNet-QA, and Video-MME benchmarks demonstrates consistent outperformance over the prior state-of-the-art VLM-RLAIF, with a +6.2% overall accuracy gain on the most rigorous test. These results validate Oracle-RLAIF as both a theoretically principled and practically effective advancement in RL-based fine-tuning for multi-modal video models."),

      divider(),

      // ─── 2. INTRODUCTION ───────────────────────────────────────────────────────
      h1("2. Introduction"),
      h2("2.1 Background and Motivation"),
      body("State-of-the-art video-language models (VLMs) are typically trained through a two-stage pipeline. In the first stage, supervised fine-tuning (SFT) on annotated video–question–answer triplets enables the model to produce syntactically coherent, contextually relevant responses. In the second stage, Reinforcement Learning from Human Feedback (RLHF) further refines the model by incorporating preference signals: human annotators compare pairs of model outputs, their judgments train a reward model (RM), and the RM's scalar scores drive policy gradient updates. This pipeline has proven essential for achieving high-quality multi-modal understanding."),
      spacer(),
      body("Nevertheless, two compounding challenges arise as VLMs scale. First, the data acquisition bottleneck: collecting reliable human preference labels for video content is both slow and expensive, precluding the feedback density that effective RL training requires. Second, the reward model fragility problem: a reward model must generalize to arbitrary prompt–response combinations, yet small distributional mismatches between training and deployment conditions introduce reward misspecification — leading the policy to exploit the RM rather than improve genuine task performance, a phenomenon known as reward hacking."),
      spacer(),
      body("Reinforcement Learning from AI Feedback (RLAIF) (Bai et al., 2022; Lee et al., 2023) addresses the data bottleneck by substituting an AI judge for human annotators. VLM-RLAIF (Ahn et al., 2024), the leading prior approach for video models, instantiates this with a dedicated 13B-parameter reward model trained on ActivityNet video narratives (over 99,000 captioned videos) to produce calibrated scalar scores for PPO-based policy optimization. While this yields significant improvements over SFT alone, the pipeline remains tightly coupled to its training domain and carries substantial data preparation costs — constraints that limit its generalizability and accessibility."),

      h2("2.2 The Oracle-RLAIF Proposal"),
      body("Oracle-RLAIF (Shi et al., 2025) re-examines the necessity of a scalar reward model from first principles. The key insight is that policy optimization does not require knowing how good a response is in absolute terms — it only requires knowing which response is better. Accordingly, Oracle-RLAIF replaces the trained RM with a drop-in Oracle ranker: any model capable of ordering N candidate responses by quality. This design is strictly weaker in its requirements than score-based RM training, yet, as the authors demonstrate, sufficient to drive effective RL alignment. The relaxed interface enables the use of general-purpose closed-source models, cross-domain AI systems, or legacy rankers as Oracle, substantially broadening the framework's applicability."),
      spacer(),
      body("To translate ordinal feedback into a principled policy gradient signal, the authors introduce GRPOrank — a rank-adapted extension of Group Relative Policy Optimization (GRPO). Where GRPO normalizes scalar rewards within a response group to compute advantages, GRPOrank replaces this with a normalized Discounted Cumulative Gain (nDCG) penalty that captures both the direction and severity of rank disagreement between the model and the Oracle. The result is a bounded, zero-sum, position-sensitive advantage function that can be directly plugged into a clipped surrogate objective."),

      h2("2.3 Scope of This Review"),
      body("This review is organized to provide an RL-grounded analysis of Oracle-RLAIF. We trace the theoretical lineage from MDPs and policy gradients through PPO and GRPO to GRPOrank, then examine each methodological component — data foundations, environment design, MDP formulation, and algorithm architecture — with explicit reference to the RL concepts they embody. The review concludes with an analysis of the empirical results through an RL-theoretic lens, identifying where rank-based optimization succeeds and where its limitations emerge."),

      divider(),

      // ─── 3. PROBLEM IDENTIFICATION ─────────────────────────────────────────────
      h1("3. Problem Identification and Solution Suitability"),
      h2("3.1 Identifying the Core Limitations"),
      body("Oracle-RLAIF targets two interconnected problems that constrain the scalability and reliability of existing RLAIF frameworks for video-language models:"),
      bullet("Reward model cost and domain coupling: VLM-RLAIF's reward model requires a dedicated pipeline — video narrative generation over 99,000+ ActivityNet clips, preference labeling, and supervised RM training — creating domain-specific coupling that cannot be easily transferred to new tasks or datasets."),
      bullet("Reward magnitude instability: PPO's advantage estimation, A_t = R_t - V_phi(s_t), depends on the absolute magnitude and consistency of scalar reward scores. Miscalibration introduces gradient instability and can cause reward hacking, where the policy learns to satisfy the RM's quirks rather than genuinely improve response quality (Chen et al., 2024; Shen et al., 2024)."),
      bullet("Value function overhead: PPO requires a separately trained value network V_phi to estimate expected return at each token state. In large language models, this is a substantial additional model to maintain, train, and stabilize."),
      bullet("Inflexibility of judge selection: Because a scalar RM must output numerically meaningful rewards, only purpose-trained domain-specific models qualify. General-purpose or closed-source models cannot serve as judges in existing frameworks."),
      spacer(),
      body("Collectively, these limitations mean that existing RLAIF pipelines are expensive, fragile, and difficult to generalize — all of which Oracle-RLAIF directly resolves."),

      h2("3.2 Why Reinforcement Learning is the Appropriate Framework"),
      body("The problem of improving generative model outputs from preference or ranking feedback is fundamentally an RL problem, and cannot be fully addressed by supervised learning alone. Supervised fine-tuning optimizes the cross-entropy loss on a fixed training distribution; it can teach the model to produce plausible responses but cannot directly optimize for which response a judge would prefer. RL provides the necessary theoretical machinery for three reasons."),
      spacer(),
      body("First, the evaluation signal arrives at the sequence level: a response is only judged after full generation, corresponding to an episodic, trajectory-level reward in RL terminology. Second, the goal is to shift the model's response distribution toward higher-quality outputs — a policy optimization problem in which the policy (the language model) must explore its action space (the token vocabulary) to discover better response strategies. Third, relative comparison of multiple candidate responses is a natural RL construct: generating G responses per query and comparing them corresponds to evaluating G trajectories per episode, which is precisely how GRPO computes within-group advantages."),
      spacer(),
      body("Oracle-RLAIF is therefore not merely applying RL as a computational tool, but exploiting the structural alignment between preference-based learning and the RL framework's core mechanisms: trajectory generation, reward assignment, and policy gradient optimization."),

      h2("3.3 Suitability of Ordinal Ranking as the Reward Signal"),
      body("The substitution of scalar scores with ordinal ranks is theoretically well-motivated. Ranking is a strictly weaker information structure than scoring — a ranker need only determine which response is better, not by how much. This relaxation offers concrete advantages within the RL pipeline:"),
      bullet("Normalization by construction: Rank values are bounded integers in {0, ..., K-1}. The nDCG transform maps these to a bounded continuous penalty in [0, 1), eliminating the reward scale sensitivity that destabilizes PPO and requires careful reward normalization preprocessing."),
      bullet("Compatibility with group-relative optimization: GRPO-family algorithms are designed around relative performance within a response group. Rank feedback is intrinsically relative, making it a natural input to GRPO's advantage estimation framework."),
      bullet("Oracle reliability: A judge model can rank responses more reliably than it can score them with calibrated magnitudes. Ordinal consistency is a weaker requirement than cardinal consistency, making the Oracle's feedback more trustworthy and reducing reward misspecification."),

      divider(),

      // ─── 4. METHODOLOGY ────────────────────────────────────────────────────────
      h1("4. Methodology"),
      h2("4.1 Data Foundation and Demand Modelling"),
      body("Oracle-RLAIF is evaluated against a controlled baseline: both Oracle-RLAIF and VLM-RLAIF begin from the identical VLM-SFT 7B checkpoint published by Ahn et al. (2024). This experimental design isolates the effect of the fine-tuning strategy from any variation in the initial model quality. The SFT checkpoint was trained on a large, curated multi-source dataset structured as a curriculum:"),
      spacer(),

      makeTable(
        ["Data Source", "Scale", "RL-Relevant Purpose"],
        [
          ["Synthetic video-text instruction tuning", "80,000 samples", "Establish baseline video-to-language alignment"],
          ["Video question-answering datasets", "67,000 samples", "Ground responses in visual reasoning tasks"],
          ["Object-centric generated datasets", "180,000 samples", "Develop object recognition and grounding"],
          ["Easy curriculum split (short answers)", "214,000 total", "Initialize stable policy for basic QA"],
          ["Hard curriculum split (long answers)", "113,000 total", "Extend policy to complex multi-hop reasoning"],
        ],
        [3200, 2400, 3760]
      ),
      spacer(),
      body("The curriculum structure — progressing from low-complexity to high-complexity tasks — reflects a core principle of RL-based training: exposing the agent to tasks of graduated difficulty stabilizes early policy learning and reduces the likelihood of converging to degenerate local optima. The difficulty metric used here (correct answer length) serves as a proxy for response complexity."),
      spacer(),
      body("For the RLAIF training phase, the Oracle ranker is trained on the same pairwise preference dataset as VLM-RLAIF's reward model, but with the ActivityNet narrative captions deliberately omitted. This design choice is theoretically significant: it decouples the Oracle from domain-specific contextual grounding, demonstrating that even a less specialized ranker suffices when the learning algorithm is designed to exploit ordinal rather than scalar feedback. The Oracle's task is strictly reduced to ordering responses — a more tractable and generalizable requirement than score generation."),

      h2("4.2 Simulation Environment Design"),
      body("In the context of language model RL, the 'environment' is not a physical or simulated world but the generative process of the policy model, evaluated through Oracle feedback. The Oracle-RLAIF training loop constitutes a complete on-policy RL environment with the following structure:"),
      spacer(),
      infoBox("Oracle-RLAIF Fine-Tuning Loop",
        "Step 1 — State Initialization: A video clip paired with a natural-language prompt constitutes the initial state s_0 for each episode.\n\nStep 2 — Trajectory Generation: The policy model pi_theta generates G = 5 complete response trajectories {o_1, ..., o_G} by autoregressively sampling from its output distribution for the same input.\n\nStep 3 — Oracle Ranking (Reward Assignment): The Oracle ranker receives all G responses along with the video-prompt context and assigns ranks {rank_1, ..., rank_G}, where rank 0 denotes the highest-quality response.\n\nStep 4 — Advantage Computation: GRPOrank computes nDCG-based penalties delta_i by comparing the policy model's predicted ranking (derived from log-probabilities) to the Oracle's ground-truth ranking, then constructs the group-relative advantage A_rank.\n\nStep 5 — Policy Update: The clipped surrogate GRPOrank objective updates the policy parameters, incorporating KL divergence and entropy regularization to maintain proximity to the reference policy and prevent distributional collapse.\n\nStep 6 — Reference Policy Refresh: At the start of each epoch, the current policy is frozen as the new reference policy pi_ref for the next epoch's KL constraint."),
      spacer(),
      body("This loop is entirely self-contained: it requires no external reward database, value network, or pre-computed reference trajectories. The Oracle serves as the sole reward signal, and on-policy generation ensures the training distribution remains aligned with the current policy's behavior throughout optimization — a key advantage over offline methods."),

      h2("4.3 MDP Formulation"),
      body("The Oracle-RLAIF fine-tuning problem admits a rigorous formalization as a token-level Markov Decision Process with sequence-level reward assignment. This formulation makes explicit how each component of the RL framework maps onto the language generation problem:"),
      spacer(),

      makeTable(
        ["MDP Component", "Oracle-RLAIF Realization"],
        [
          ["State (s_t)", "Full generation context at step t: video frames + prompt + partial response (o_{i,<t})"],
          ["Action (a_t)", "Next token o_{i,t} selected from the vocabulary V under the current policy"],
          ["Policy (pi_theta)", "The VLM's conditional token distribution: pi_theta(o_{i,t} | q, o_{i,<t})"],
          ["Reward (R)", "Episodic, sequence-level: Oracle rank_i is converted to nDCG-based advantage at episode end"],
          ["Transition (T)", "Deterministic: T(s_t, a_t) = s_{t+1} = (video, prompt, o_{i,1}, ..., o_{i,t})"],
          ["Episode", "Complete response generation from first token to end-of-sequence (EOS) token"],
          ["Discount (gamma)", "Effectively 1.0 within an episode — rank-based advantage is broadcast uniformly to all tokens"],
          ["Objective", "max_theta E_{q~D, {o_i}~pi_theta} [sum_i A_rank_i], subject to KL(pi_theta || pi_ref) <= delta"],
        ],
        [3000, 6360]
      ),
      spacer(),
      body("A critical RL design choice is credit assignment: because the Oracle's rank is assigned to the complete response rather than individual tokens, the advantage A_rank is broadcast uniformly across all tokens in the response. This is standard in LLM RL fine-tuning (consistent with GRPO and PPO for language models) and reflects the assumption that all generation decisions within a response contribute equally to the final quality judgment."),
      spacer(),
      body("The KL divergence constraint between the current policy pi_theta and the frozen reference policy pi_ref plays the role of a trust region in the RL sense: it bounds the step size of policy updates, preventing catastrophic forgetting of SFT-learned behaviors and stabilizing the training dynamics. The entropy bonus H[pi_theta] further guards against premature policy collapse — the degenerate mode-seeking behavior where the policy generates the same response repeatedly."),

      h2("4.4 Algorithm Selection and Architecture"),
      h3("4.4.1 From PPO to GRPO: Eliminating the Value Function"),
      body("Proximal Policy Optimization (PPO) underpins VLM-RLAIF's RL phase. PPO stabilizes policy gradient updates by clipping the importance ratio r_t(theta) = pi_theta(a_t|s_t) / pi_theta_old(a_t|s_t) within a trust region, and computes advantage as A_t = R_hat_t - V_phi(s_t), where V_phi is a separately trained value function approximating expected cumulative reward. While effective for scalar reward settings, PPO's reliance on V_phi introduces two compounding difficulties in Oracle-RLAIF's setting: (1) maintaining and training a separate value network is computationally expensive at the 7B parameter scale, and (2) when the reward signal is ordinal rather than cardinal, raw rank values carry no absolute magnitude — rank 0 is better than rank 1, but the difference is not quantitatively comparable to the difference between ranks 3 and 4. Feeding such values into a value-function estimator produces semantically meaningless advantage estimates."),
      spacer(),
      body("Group Relative Policy Optimization (GRPO), introduced by DeepSeek for mathematical reasoning (Shao et al., 2024), resolves both issues by eliminating the value function entirely. GRPO generates G candidate responses per query and replaces V_phi with the group's own mean reward as the baseline:"),
      spacer(),
      infoBox("GRPO Advantage Formula",
        "A_{i,t} = [R(q, o_i) - mu_R(q)] / sigma_R(q)\n\nWhere R(q, o_i) is the scalar reward for response i, mu_R(q) is the mean reward across all G responses in the group, and sigma_R(q) is their standard deviation. This z-score normalization achieves three properties simultaneously: (1) advantages are automatically zero-centered within each group, (2) the learning signal is invariant to the absolute scale and shift of reward values, and (3) no separate value model or critic network is required — the group itself serves as the internal baseline."),
      spacer(),
      body("This group-relative structure is precisely why GRPO is the appropriate foundation for Oracle-RLAIF. Since ranks are inherently relative — response A is better than response B only in the context of the group — a group-relative advantage formulation captures the full semantic content of rank feedback without requiring scalar magnitude."),

      h3("4.4.2 GRPOrank: Rank-Adapted Policy Gradient"),
      body("GRPOrank extends GRPO by replacing the z-score normalized reward advantage with a rank-derived advantage grounded in normalized Discounted Cumulative Gain (nDCG) — a standard metric from information retrieval's learning-to-rank literature. The key design decisions and their RL justifications are as follows."),
      spacer(),
      body("The penalty for response i is defined as delta_i = 1 - nDCG_i = 1 - DCG(rank_hat_i) / DCG(rank_i), where rank_i is the Oracle's ground-truth rank, rank_hat_i is the policy's predicted rank (obtained by sorting responses in descending order of their log-probability under the current policy), and DCG(rank) = 1 / [(1 + rank) * log_2(2 + rank)]. The advantage function is then: A_rank = E_{j in G_i}[delta_j] - delta_i, i.e., the group mean penalty minus the individual penalty. Responses with penalty below the group average receive positive advantage (the policy is encouraged to assign them higher probability), and those above receive negative advantage."),
      spacer(),

      makeTable(
        ["True Rank", "Predicted Rank", "nDCG", "Penalty (delta)", "Advantage (K=5)"],
        [
          ["0 (best)", "0 — correct", "1.0000", "0.0000", "+0.2887 (promoted)"],
          ["0 (best)", "1 — off by one", "0.7925", "0.2075", "+0.0812"],
          ["0 (best)", "2", "0.6667", "0.3333", "-0.0446"],
          ["0 (best)", "4 — ranked worst", "0.5170", "0.4830", "-0.1943 (max penalty)"],
          ["4 (worst)", "3 — off by one", "0.9757", "0.0243", "~0 (negligible)"],
        ],
        [1560, 1800, 1440, 1800, 2760]
      ),
      spacer(),
      body("The position-sensitivity of GRPOrank is evident in the table: misidentifying the best response (true rank 0) as worst (predicted rank 4) produces the maximum penalty of delta = 0.4830, while a minor misranking of the lowest-quality response incurs a negligible delta = 0.0243. This asymmetry encodes a key RL design principle: the learning signal should prioritize the outcomes that matter most to end-users — getting the top-ranked response right."),

      h3("4.4.3 Mathematical Properties of GRPOrank"),
      body("The GRPOrank advantage function satisfies three formally provable properties that make it a principled RL training signal:"),
      spacer(),
      bodyMixed([
        { text: "Zero-Sum Property: ", bold: true },
        { text: "The sum of advantages across all G responses in a group equals exactly zero: sum_i A_rank_i = 0. This guarantees that GRPOrank implements a strictly competitive relative signal — the policy update always involves simultaneously promoting some responses and suppressing others, preventing the degenerate solution where all token probabilities drift uniformly in one direction." }
      ]),
      spacer(),
      bodyMixed([
        { text: "Boundedness: ", bold: true },
        { text: "Since rank_i in {0, ..., K-1} and DCG is monotonically decreasing, nDCG_i in (0, 1] for all configurations, yielding delta_i in [0, 1). This guarantees numerically stable advantage values regardless of group size K, without requiring reward clipping or normalization preprocessing." }
      ]),
      spacer(),
      bodyMixed([
        { text: "Position-Sensitive Discounting: ", bold: true },
        { text: "The logarithmic denominator in DCG ensures that rank errors at the top of the list incur exponentially larger penalties than equivalent errors at the bottom. This directly aligns the optimization objective with deployment priorities: the model is most strongly trained to correctly identify and generate the highest-quality response." }
      ]),

      h3("4.4.4 Model Architecture"),
      body("The policy VLM in Oracle-RLAIF inherits the architecture from VLM-RLAIF's SFT checkpoint, comprising three tightly integrated components:"),
      bullet("Language Backbone: LLaMA-2-7B — a 7-billion parameter autoregressive causal transformer that generates token-by-token responses conditioned on multimodal context"),
      bullet("Vision Encoder: Frozen OpenAI CLIP ViT-Large/14-336 — converts uniformly sampled video frames (50 per video) into spatial-temporal feature representations aligned to the language embedding space"),
      bullet("Vision-Language Adapter: Q-Former with 32 learnable query tokens, followed by two linear projection layers, bridging the fixed-dimension vision features to the language model's input space"),
      bullet("Efficient RL Fine-Tuning: QLoRA (Dettmers et al., 2023) with 4-bit quantization and low-rank adapters (rank=65, alpha=16), enabling RL training on 4x NVIDIA H100 GPUs — a substantial infrastructure reduction over full-precision fine-tuning"),
      spacer(),
      body("The Oracle ranker shares architectural lineage with the policy model but operates in evaluation mode: given a video, prompt, and G candidate responses, it produces a complete ordering. Critically, Oracle-RLAIF's Oracle was trained without domain-specific narrative data, confirming that the framework's performance advantage stems from the GRPOrank optimization strategy rather than superior Oracle quality."),

      divider(),

      // ─── 5. EMPIRICAL RESULTS ──────────────────────────────────────────────────
      h1("5. Empirical Results and Analysis"),
      h2("5.1 Benchmark Performance"),
      body("The evaluation protocol is structured in two experimental regimes, chosen to isolate different dimensions of the fine-tuning comparison. In the first regime, both models are evaluated on three open-ended video QA benchmarks (MSVD-QA, MSRVTT-QA, ActivityNet-QA), with responses judged by GPT-3.5-turbo on five quality dimensions: relevance, detail capture, contextual understanding, temporal reasoning, and consistency. Both models start from the same VLM-SFT 7B checkpoint, making any performance differential attributable solely to the RL fine-tuning strategy:"),
      spacer(),

      makeTable(
        ["Benchmark", "VLM-RLAIF Acc.", "Oracle-RLAIF Acc.", "Gain"],
        [
          ["MSVD-QA", "68.5%", "72.9%", "+4.4%"],
          ["MSRVTT-QA", "54.2%", "59.2%", "+5.0%"],
          ["ActivityNet-QA", "46.1%", "48.1%", "+2.0%"],
        ],
        [2500, 2400, 2400, 2060]
      ),
      spacer(),
      body("The second and more rigorous evaluation used Video-MME — a multiple-choice benchmark explicitly designed to preclude training data leakage. Oracle-RLAIF achieved a +6.2% overall accuracy improvement, with the most pronounced gains in Temporal Perception (+21.2%), Action Recognition (+11.7%), and Object Reasoning (+11.2%). These categories share a common structure: they require the model to recognize causally ordered events and ground responses in temporally specific visual evidence — precisely the type of discrimination task for which clear, consistent Oracle ranking is achievable."),
      spacer(),
      body("Performance declined relative to VLM-RLAIF in Spatial Perception (-2.6%) and Spatial Reasoning (-3.8%). The authors attribute this to the higher ambiguity of spatial and abstract quality distinctions, which makes Oracle rankings less reliable in these categories. Under GRPOrank, noisy rankings propagate as noisy advantage estimates into the policy gradient, potentially destabilizing training in the very directions where the Oracle's signal is least trustworthy."),

      h2("5.2 RL-Theoretic Interpretation of Results"),
      body("The observed performance pattern is theoretically predictable from RL first principles. The fundamental determinant of policy gradient learning quality is the reliability of the reward signal: when the Oracle consistently ranks higher-quality responses above lower-quality ones, GRPOrank's advantage function is well-calibrated, and the policy gradient reliably pushes the model toward better responses. Temporal and action-based tasks fulfill this condition — the quality difference between a response that correctly identifies an event sequence and one that does not is objectively measurable and ranks are consistent."),
      spacer(),
      body("Spatial reasoning and abstract inference tasks violate this condition. When two responses are both plausible descriptions of the same spatial configuration, the Oracle's ranking is effectively noisy. Noise in the reward function translates directly to variance in the policy gradient estimate, increasing the risk of destructive updates. This is a well-known challenge in RL: high reward variance requires either variance reduction techniques (such as better baselines) or more robust advantage estimation — both directions the authors identify as avenues for future work."),
      spacer(),
      body("This analysis demonstrates that Oracle-RLAIF's performance profile is not merely an empirical observation but a direct consequence of its RL-theoretic design: the framework performs best where its core assumption — that ranking quality is reliable and consistent — is most satisfied."),

      divider(),

      // ─── 6. RELATION TO BROADER RL LITERATURE ─────────────────────────────────
      h1("6. Relation to Broader RL Literature"),
      body("Oracle-RLAIF occupies a well-defined position at the convergence of several active research directions in reinforcement learning for language model alignment. Understanding its relationship to prior work clarifies both what makes it novel and what theoretical lineage it inherits."),
      spacer(),
      bodyMixed([{ text: "RLHF and RLAIF: ", bold: true }, { text: "Ouyang et al. (2022) established the modern RLHF paradigm, in which human preference labels train a reward model whose scalar outputs drive PPO. Bai et al. (2022) and Lee et al. (2023) introduced RLAIF by substituting an AI judge for human annotators, maintaining the same scalar reward structure. Oracle-RLAIF departs from this lineage by removing the requirement of a calibrated scalar reward model altogether, replacing it with a more accessible ordinal interface." }]),
      spacer(),
      bodyMixed([{ text: "GRPO and DeepSeek-R1: ", bold: true }, { text: "GRPO (Shao et al., 2024) demonstrated that group-relative advantage estimation is a highly effective replacement for PPO's value-function-based advantage in language model fine-tuning, contributing to DeepSeek-R1's strong reasoning performance. Oracle-RLAIF inherits GRPO's value-free, group-relative structure and extends it to handle ordinal rather than cardinal reward signals — a conceptually natural extension that the original GRPO formulation does not support." }]),
      spacer(),
      bodyMixed([{ text: "Listwise Learning to Rank: ", bold: true }, { text: "The nDCG metric is a cornerstone of the learning-to-rank literature in information retrieval, where optimizing ranked list quality over documents or passages is the core task. Oracle-RLAIF is the first work, to the authors' knowledge, to integrate a listwise ranking metric directly into a policy gradient objective, bridging the information retrieval and RL communities in a principled way." }]),
      spacer(),
      bodyMixed([{ text: "DPO and Offline Preference Optimization: ", bold: true }, { text: "Direct Preference Optimization (Rafailov et al., 2023) and related approaches (RRHF, RAFT) align language models with preference data without an explicit reward model by recasting the RLHF objective as a supervised loss over preference pairs. While these methods simplify training infrastructure, they operate in an offline or supervised regime and cannot exploit the on-policy exploration benefits of RL. Oracle-RLAIF is a fully online RL method — generating and evaluating new trajectories at every update step — which enables continuous adaptation of the training distribution to the policy's current behavior." }]),

      divider(),

      // ─── 7. CONCLUSION ─────────────────────────────────────────────────────────
      h1("7. Conclusion"),
      body("Oracle-RLAIF constitutes a theoretically grounded and empirically validated advancement in reinforcement learning for multi-modal video language model alignment. Its central contribution is the decoupling of the RLAIF pipeline from the requirement of a domain-specific, calibrated scalar reward model — achieved by substituting an Oracle ranker and a rank-adapted policy gradient algorithm, GRPOrank. This redesign simultaneously reduces infrastructure costs, broadens the range of applicable Oracle models, and, as demonstrated by the empirical results, improves video comprehension performance across multiple benchmarks."),
      spacer(),
      body("For a student of reinforcement learning, Oracle-RLAIF offers a particularly instructive case study in applied RL design. Each algorithmic choice — the episodic MDP structure with sequence-level reward, the shift from PPO's value-function-based advantage to GRPO's group-relative normalization, and the further extension to nDCG-based rank advantage — is motivated by a specific deficiency in the preceding approach and traceable to a well-understood RL principle. The paper's evolution from PPO through GRPO to GRPOrank exemplifies how RL methodology adapts when the reward signal structure changes: from cardinal to ordinal, from absolute to relative, and from individual to group-comparative."),
      spacer(),
      body("The identified performance limits — weaker gains on spatial and abstract tasks where Oracle rankings are inherently noisier — highlight a fundamental principle that applies across all RL systems: the ceiling on learned policy quality is set by the quality and consistency of the reward signal. Future work expanding GRPOrank to settings with richer multimodal Oracle models, or combining rank signals with auxiliary spatial encoders, represents a natural progression building directly on this analysis."),
      spacer(),

      infoBox("Key RL Concepts Illustrated by This Paper",
        "1. MDP Formulation: Language generation is a token-level MDP with sequence-level, episodic reward — the Oracle's rank is assigned to the complete response trajectory, not individual token actions.\n\n2. Policy Gradient Optimization: GRPOrank is a clipped surrogate policy gradient method — it maximizes expected advantage while constraining the update magnitude via importance ratio clipping.\n\n3. Value-Free Advantage Estimation: GRPO eliminates the critic network by using within-group mean reward as the baseline, making advantage estimation self-contained and robust to reward scale variation.\n\n4. Ordinal-to-Continuous Reward Transformation: nDCG maps integer ranks to a continuous, bounded [0,1) penalty, enabling differentiation-compatible advantage computation from inherently discrete ordinal feedback.\n\n5. Trust Region and Entropy Regularization: KL divergence from the reference policy bounds the policy update step size; entropy bonus prevents mode collapse — both standard RL stability mechanisms applied to the language generation setting.\n\n6. Reward Signal Quality as a Performance Determinant: The gap between Oracle-RLAIF's gains on temporal tasks (+21.2%) and losses on spatial tasks (-2.6%) directly reflects the reliability of the Oracle's rankings in each task category — a textbook consequence of the dependence of policy gradient variance on reward signal noise."),

      divider(),

      // ─── REFERENCES ─────────────────────────────────────────────────────────────
      h1("References"),
      body("Ahn, D., Hu, Y., Ostapenko, O., et al. (2024). Tuning large multimodal models for videos using reinforcement learning from AI feedback. ACL 2024."),
      body("Bai, Y., Jones, A., Ndousse, K., et al. (2022). Training a helpful and harmless assistant with reinforcement learning from human feedback. arXiv:2204.05862."),
      body("DeepSeek-AI. (2025). DeepSeek-R1: Incentivizing reasoning capability in LLMs via reinforcement learning. arXiv:2501.12948."),
      body("Fu, C., Dai, Y., Luo, Y., et al. (2025). Video-MME: The first-ever comprehensive evaluation benchmark of multi-modal LLMs in video analysis. CVPR 2025."),
      body("Lee, H., Phatale, S., Mansoor, H., et al. (2023). RLAIF: Scaling reinforcement learning from human feedback with AI feedback. arXiv:2309.00267."),
      body("Ouyang, L., Wu, J., Jiang, X., et al. (2022). Training language models to follow instructions with human feedback. NeurIPS 2022."),
      body("Rafailov, R., Sharma, A., Mitchell, E., et al. (2023). Direct preference optimization: Your language model is secretly a reward model. NeurIPS 2023."),
      body("Schulman, J., Wolski, F., Dhariwal, P., Radford, A., and Klimov, O. (2017). Proximal policy optimization algorithms. arXiv:1707.06347."),
      body("Shao, Z., Wang, P., Zhu, Q., et al. (2024). DeepSeekMath: Pushing the limits of mathematical reasoning in open language models. arXiv:2402.03300."),
      body("Shi, D., Glatt, R., Klymko, C., et al. (2025). Oracle-RLAIF: An improved fine-tuning framework for multi-modal video models through reinforcement learning from oracle ranking feedback. ARLET Workshop, arXiv:2510.02561."),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/Oracle_RLAIF_Literature_Review.docx', buffer);
  console.log('Done');
});
