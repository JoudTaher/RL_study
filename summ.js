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
      body("This literature review examines Oracle-RLAIF (Reinforcement Learning from AI Feedback through Oracle Preferences), a novel framework introduced by Shi et al. (2025) for fine-tuning large video-language models (VLMs). The paper addresses a core bottleneck in modern AI alignment pipelines: the prohibitive cost and complexity of reward model training when scaling VLMs with reinforcement learning. The central RL contribution of the work is twofold. First, the authors replace the conventional trained reward model with a drop-in Oracle ranker — a general-purpose model that produces ordinal rankings rather than scalar scores over candidate responses. Second, they introduce GRPOrank, a rank-adapted variant of Group Relative Policy Optimization (GRPO) that directly incorporates ordinal feedback into the policy gradient update via a normalized Discounted Cumulative Gain (nDCG) penalty."),
      spacer(),
      body("This review traces the RL foundations required to understand the paper, including Markov Decision Processes (MDPs), policy optimization, Proximal Policy Optimization (PPO), GRPO, and the transition from RLHF to RLAIF in language model alignment. We analyze the paper's problem framing, solution design, experimental methodology, and empirical results across standard video QA benchmarks (MSVD, MSRVTT, ActivityNet) and the more rigorous Video-MME dataset. Our analysis confirms that Oracle-RLAIF represents a meaningful and well-motivated advancement in reinforcement learning applied to multi-modal AI, with the rank-based optimization demonstrating particular strength in temporally and causally grounded tasks."),

      divider(),

      // ─── 2. INTRODUCTION ───────────────────────────────────────────────────────
      h1("2. Introduction"),
      h2("2.1 Background and Motivation"),
      body("Modern large language models and video-language models achieve state-of-the-art performance through a two-stage training process: supervised fine-tuning (SFT) followed by reinforcement learning from human feedback (RLHF). SFT teaches the model to produce syntactically correct and contextually relevant outputs by training on annotated data. RLHF then further refines the model's behavior by incorporating human preference signals — humans label which of two model responses is better, this preference data trains a reward model, and the reward model's scores drive RL-based policy updates."),
      spacer(),
      body("However, as models scale to billions of parameters and video understanding demands richer temporal and spatial comprehension, two major challenges emerge. First, collecting high-quality human preference labels at scale is extremely expensive and slow, creating a significant bottleneck. Second, training a reliable reward model that generalizes across arbitrary prompt-response combinations is technically difficult — small misalignments in the reward signal can cause the policy to exploit the reward model rather than genuinely improve."),
      spacer(),
      body("These challenges have motivated the development of Reinforcement Learning from AI Feedback (RLAIF), where an AI model acts as the judge in place of a human. The most advanced prior approach for VLMs, called VLM-RLAIF (Ahn et al., 2024), uses a dedicated 13-billion-parameter reward model trained on video narrative data to score candidate responses. While effective, this approach inherits significant costs: training that reward model requires a specialized pipeline including video captioning over 99,000 videos with ActivityNet data. The resulting system is tightly coupled to a specific dataset and domain, limiting its generalizability."),

      h2("2.2 The Oracle-RLAIF Proposal"),
      body("Oracle-RLAIF decouples the RLAIF framework from the requirement of a fully calibrated, score-generating reward model. Instead, the framework asks only for an Oracle model capable of producing a relative ranking of N candidate responses generated by the policy VLM for a given prompt. This is a substantially easier and more flexible requirement: ranking is cognitively and computationally simpler than absolute scoring, and it opens the door to using closed-source commercial models, legacy AI systems, or general-purpose multimodal models as the Oracle."),
      spacer(),
      body("Alongside this framework redesign, the authors introduce GRPOrank — a modification of GRPO that replaces scalar reward-based advantage estimation with a rank-based advantage derived from nDCG penalties. This ensures the training signal directly encodes ordinal information, with position-sensitive discounting that penalizes misranking high-quality responses more than misranking low-quality ones."),

      h2("2.3 Purpose of This Review"),
      body("This literature review is structured to serve a student of reinforcement learning who requires both conceptual grounding and technical depth. We begin with the RL foundations embedded in the paper, then move systematically through the paper's methodology: its problem formulation, solution architecture, algorithm design, and empirical validation. Each section links the paper's specific choices back to fundamental RL principles, making the work accessible and analytically complete."),

      divider(),

      // ─── 3. PROBLEM IDENTIFICATION ─────────────────────────────────────────────
      h1("3. Problem Identification and Solution Suitability"),
      h2("3.1 The Core Problem"),
      body("The primary problem addressed by Oracle-RLAIF is the cost, fragility, and inflexibility of reward model training in RLAIF pipelines for video-language models. Specifically:"),
      bullet("Existing RLAIF frameworks require a trained reward model capable of assigning calibrated scalar scores to arbitrary prompt-response pairs — a technically demanding requirement."),
      bullet("Training this reward model in VLM-RLAIF required video narrative generation via ActivityNet captions (99,000+ videos), creating domain-specific coupling."),
      bullet("Scalar reward models are known to suffer from reward hacking — the policy can learn to exploit the reward model's weaknesses rather than improving genuine video comprehension."),
      bullet("As VLMs grow in size, the cost of gathering preference data and maintaining specialized reward models grows proportionally, limiting scalability."),
      spacer(),
      body("A secondary technical problem concerns the optimization algorithm. PPO, used in VLM-RLAIF, requires a trained value function (critic) to estimate advantage — adding memory and computational overhead. Furthermore, PPO's advantage estimation depends on reward magnitude, making it sensitive to the scale and calibration of the reward model's outputs."),

      h2("3.2 Why Reinforcement Learning is the Right Tool"),
      body("Reinforcement learning is appropriate here for several fundamental reasons. The task of improving response quality from preference or ranking feedback is inherently sequential and non-differentiable with respect to the model's token generation process. Standard supervised learning with cross-entropy loss cannot directly optimize for which response a judge would prefer — it can only mimic the distribution of training tokens. RL provides a principled framework for optimizing non-differentiable objectives (such as human preference rankings) by treating the language model as a policy, each generated response as a trajectory, and the Oracle's ranking as a reward signal."),
      spacer(),
      body("More specifically, the key features of the problem that make RL suitable are: (1) the evaluation signal (ranking) arrives after a complete response is generated, analogous to episodic RL reward at trajectory end; (2) we want to shift the distribution of responses toward higher-quality outputs, which is a policy optimization problem; and (3) relative comparison of multiple responses (which underpins GRPO and GRPOrank) naturally fits the RL setting of sampling and evaluating trajectories."),

      h2("3.3 Why Oracle Ranking is Better than Score-Based Reward"),
      body("Replacing scalar reward scoring with ordinal ranking offers several advantages that are well-grounded in RL theory and practice:"),
      bullet("Ranking is a weaker requirement than scoring: it does not demand that the Oracle assign numerically meaningful magnitudes to quality differences, only relative ordering. This makes the Oracle more reliable and general."),
      bullet("Rankings are naturally normalized: because rank values are bounded integers (0 to K-1), the resulting advantage estimates are automatically bounded, avoiding the reward magnitude instability that affects PPO with raw scalar rewards."),
      bullet("The GRPO family of algorithms, which Oracle-RLAIF builds upon, is specifically designed to extract learning signal from relative comparisons within a group of responses — making it the natural algorithmic match for rank-based feedback."),
      bullet("Ordinal feedback is more robust to Oracle calibration errors: a judge can reliably say response A is better than response B without needing to precisely quantify the degree of superiority."),

      divider(),

      // ─── 4. METHODOLOGY ────────────────────────────────────────────────────────
      h1("4. Methodology"),
      h2("4.1 Data Foundation and Demand Modelling"),
      body("Oracle-RLAIF begins from the same pretrained supervised fine-tuned (SFT) checkpoint as VLM-RLAIF, specifically the VLM-SFT 7B model published by Ahn et al. (2024). This ensures that any performance differences between Oracle-RLAIF and VLM-RLAIF are attributable to the fine-tuning strategy rather than the initial model quality. The SFT model was itself trained on a large and carefully curated dataset:"),
      spacer(),

      makeTable(
        ["Data Source", "Size", "Purpose"],
        [
          ["Synthetic video-text instruction tuning", "80,000 samples", "Teach basic video-to-text alignment"],
          ["Video question answering datasets", "67,000 samples", "Ground responses in visual reasoning"],
          ["Object-centric generated datasets", "180,000 samples", "Enhance object recognition and description"],
          ["Easy curriculum split", "214,000 total", "Develop basic QA capability first"],
          ["Hard curriculum split", "113,000 total", "Progress to complex comprehension tasks"],
        ],
        [3200, 2400, 3760]
      ),
      spacer(),
      body("The curriculum learning strategy — ordering training from easy to hard based on correct answer length — reflects an important principle in RL-inspired training: presenting the model with tasks of graduated difficulty accelerates learning and prevents early convergence to poor local optima."),
      spacer(),
      body("For the RLAIF phase, the Oracle ranker in Oracle-RLAIF is trained on the same preference data as the VLM-RLAIF reward model, but critically without the ActivityNet video narrative captions. This design choice simulates a more realistic scenario where the Oracle is a drop-in model without domain-specific narrative tuning — demonstrating that Oracle-RLAIF is more data-efficient and domain-agnostic."),

      h2("4.2 Simulation Environment Design"),
      body("Unlike classical RL environments (e.g., Atari games or robotic simulations), the 'environment' in Oracle-RLAIF is the language generation process of the VLM itself, evaluated by the Oracle ranker. The simulation loop operates as follows:"),
      spacer(),
      infoBox("Oracle-RLAIF Fine-Tuning Loop (Simulation Environment)",
        "Step 1 — Input: A video clip paired with a natural language question or prompt is fed to the policy VLM.\n\nStep 2 — Response Generation: The policy model generates N = 5 candidate responses by sampling from its output distribution for the same input.\n\nStep 3 — Oracle Ranking: The Oracle ranker receives all N responses along with the original video-prompt context and produces a ranking ordering the responses from best to worst (rank 0 = best).\n\nStep 4 — Advantage Computation: GRPOrank computes nDCG-based advantage values by comparing the model's predicted ranking (derived from log-probabilities) to the Oracle's ground-truth ranking.\n\nStep 5 — Policy Update: The GRPOrank loss function updates the policy model's weights to increase the probability of generating higher-ranked responses, with KL and entropy regularization for stability.\n\nStep 6 — Iteration: Steps 1-5 repeat across training epochs, with the reference policy frozen at the start of each epoch for importance sampling."),
      spacer(),
      body("This loop is self-contained and does not require an external environment simulator, reward database, or value network. The Oracle serves as the sole evaluator, and the policy model's own generative distribution defines the action space."),

      h2("4.3 MDP Formulation"),
      body("To understand Oracle-RLAIF from an RL perspective, we can formally characterize the learning problem as a Markov Decision Process (MDP). In language model RL, the MDP is defined at the token level, but the reward is assigned at the sequence (response) level:"),
      spacer(),

      makeTable(
        ["MDP Component", "Oracle-RLAIF Realization"],
        [
          ["State (s)", "Current token generation context: video frames + prompt + tokens generated so far (o_{i,<t})"],
          ["Action (a)", "Next token selected from the vocabulary by the policy model (o_{i,t})"],
          ["Policy (π_θ)", "The VLM's token generation probability distribution: π_θ(o_{i,t} | q, o_{i,<t})"],
          ["Reward (R)", "Derived from Oracle ranking: Oracle assigns rank_i to response i; this rank determines the advantage"],
          ["Transition", "Deterministic: appending the selected token extends the context to the next state"],
          ["Episode", "One complete response generation: from first token to EOS (end-of-sequence) token"],
          ["Objective", "Maximize expected advantage of generated responses relative to Oracle rankings"],
        ],
        [3200, 6160]
      ),
      spacer(),
      body("A crucial subtlety is that the reward is sparse and episodic — it is only assigned at the end of the full response, not at each token step. This is standard in LLM fine-tuning RL and necessitates credit assignment across the entire response. GRPOrank handles this by assigning the same advantage value A_rank to all tokens within a response, effectively broadcasting the sequence-level rank signal to every generation step."),
      spacer(),
      body("The MDP also incorporates a KL divergence constraint between the current policy and a frozen reference policy. This constraint is standard in LLM RL and serves the role of a trust region — ensuring the policy does not deviate too far from its SFT-initialized starting point, which could cause catastrophic forgetting or degenerate outputs."),

      h2("4.4 Algorithm Selection and Architecture"),
      h3("4.4.1 From PPO to GRPO: Why the Change?"),
      body("Proximal Policy Optimization (PPO) is the foundational algorithm in the VLM-RLAIF framework. PPO's loss function clips the policy importance ratio to prevent large updates and uses a learned value function (critic) to compute advantage estimates. For language models, this means maintaining a separate value head that estimates the expected cumulative reward from each state — adding memory and computational cost."),
      spacer(),
      body("PPO's advantage function is: A_t = R_t_hat - V_φ(s_t), where V_φ(s_t) is the value function approximation. When rewards come from a well-calibrated scalar reward model, this works well. However, when the reward signal is ordinal (a rank), the raw rank values carry no meaningful magnitude — rank 1 is better than rank 2, but the difference between ranks 1 and 2 is not necessarily the same as between ranks 2 and 3. Feeding raw ranks into PPO's value-based advantage would produce unstable and semantically incorrect training signals."),
      spacer(),
      body("Group Relative Policy Optimization (GRPO), introduced by DeepSeek for mathematical reasoning tasks, elegantly addresses this by eliminating the value function entirely. Instead of estimating absolute expected reward, GRPO generates G candidate responses per query and uses their relative rewards to compute normalized advantages:"),
      spacer(),
      infoBox("GRPO Advantage Formula",
        "A_{i,t} = [R(q, o_i) - μ_R(q)] / σ_R(q)\n\nWhere R(q, o_i) is the reward for response i, μ_R(q) is the mean reward across all G responses to query q, and σ_R(q) is the standard deviation. This z-score normalization means: (1) advantages are automatically zero-centered within each group, (2) the training signal is invariant to the absolute scale of rewards, and (3) no separate value model is required."),
      spacer(),
      body("This normalization is what makes GRPO ideal as the foundation for Oracle-RLAIF: since we are replacing rewards with ranks, we need an advantage formulation that is robust to the ordinal (rather than cardinal) nature of the feedback."),

      h3("4.4.2 GRPOrank: The Core Innovation"),
      body("GRPOrank extends GRPO by replacing the normalized reward advantage with a rank-derived advantage based on normalized Discounted Cumulative Gain (nDCG). This is the paper's primary algorithmic contribution and encodes three key RL-design intuitions:"),
      spacer(),
      body("Intuition 1 — Relative Comparison: Like GRPO, GRPOrank compares each response against the group average rather than an absolute standard. Responses that the policy ranks higher than the Oracle expects receive negative advantage (penalty), while responses ranked correctly receive positive advantage."),
      spacer(),
      body("Intuition 2 — Position-Sensitive Discounting: The nDCG penalty uses a logarithmic discounting function DCG(rank) = 1 / [(1 + rank) × log_2(2 + rank)]. This means errors involving the top-ranked response are penalized much more severely than errors involving low-ranked responses. This is justified because users primarily see the best response — getting the top ranking right matters most."),
      spacer(),
      body("Intuition 3 — Bounded and Stable Advantages: Because nDCG values lie in (0, 1], the penalty δ_i = 1 - nDCGi lies in [0, 1). This guarantees bounded, numerically stable advantage values with no need for reward normalization preprocessing."),
      spacer(),

      makeTable(
        ["True Rank", "Predicted Rank", "nDCG", "Penalty (δ)", "Advantage"],
        [
          ["0 (best)", "0 (correct)", "1.0000", "0.0000", "+0.2887 (rewarded)"],
          ["0 (best)", "1 (off by 1)", "0.7925", "0.2075", "+0.0812"],
          ["0 (best)", "2", "0.6667", "0.3333", "−0.0446"],
          ["0 (best)", "4 (worst)", "0.5170", "0.4830", "−0.1943 (penalized most)"],
          ["4 (worst)", "3 (off by 1)", "0.9757", "0.0243", "~0 (minor penalty)"],
        ],
        [1560, 1800, 1440, 1800, 2760]
      ),
      spacer(),
      body("The table above (reproduced from the paper's Appendix A.2 for K=5 groups) illustrates the position-sensitivity of GRPOrank. A predicted rank of 4 when the true rank is 0 (incorrectly ranking the best response as worst) yields the maximum penalty of δ = 0.4830, while mispredicting a low-ranked response's exact position incurs a negligible penalty of δ = 0.0243. This asymmetry directly encodes the user-facing priority of getting top responses right."),

      h3("4.4.3 Mathematical Properties of GRPOrank"),
      body("The GRPOrank objective possesses three mathematically guaranteed properties that make it principled and reliable as an RL training signal:"),
      spacer(),
      bodyMixed([
        { text: "Zero-Sum Property: ", bold: true },
        { text: "The sum of advantages across all responses in a group equals zero: Σ A_rank = 0. This means GRPOrank is a purely relative learning signal — some responses go up in probability, others go down, in exact balance. This property prevents degenerate solutions where all responses' probabilities shift uniformly." }
      ]),
      spacer(),
      bodyMixed([
        { text: "Boundedness: ", bold: true },
        { text: "Since rank values lie in {0, ..., K-1} and DCG is a monotonically decreasing function of rank, nDCG ∈ (0, 1] and therefore δ_i ∈ [0, 1). This guarantees that advantage values are always numerically stable and interpretable, regardless of group size K." }
      ]),
      spacer(),
      bodyMixed([
        { text: "Position-Sensitive Discounting: ", bold: true },
        { text: "The logarithmic discount in the DCG formula ensures that the learning signal is not uniform across all rank positions — it explicitly prioritizes accuracy at the top of the ranking, mirroring real-world deployment priorities." }
      ]),

      h3("4.4.4 Model Architecture"),
      body("The policy model (VLM) in Oracle-RLAIF inherits the architecture of VLM-RLAIF's SFT checkpoint:"),
      bullet("Language Backbone: LLaMA-2-7B — a 7-billion parameter autoregressive transformer"),
      bullet("Vision Encoder: Frozen OpenAI CLIP ViT-Large (patch size 14, resolution 336) — extracts spatial and temporal visual features from video frames"),
      bullet("Adapter: Q-FORMER adapter with 32 learnable tokens — bridges vision and language representations"),
      bullet("Efficient Fine-Tuning: QLoRA (Quantized Low-Rank Adaptation) — combines 4-bit quantization with low-rank adapter matrices, enabling fine-tuning on 4×NVIDIA H100 GPUs rather than requiring full-precision training on much larger infrastructure"),
      spacer(),
      body("The Oracle ranker is architecturally similar but is prompted differently — rather than generating a response, it receives the video, prompt, and N candidate responses, and outputs a ranking. Crucially, Oracle-RLAIF's Oracle ranker was trained without video narrative caption data, demonstrating that the framework succeeds even with a less specialized judge."),

      divider(),

      // ─── 5. EMPIRICAL RESULTS ──────────────────────────────────────────────────
      h1("5. Empirical Results and Analysis"),
      h2("5.1 Benchmark Performance"),
      body("Oracle-RLAIF was evaluated on two experimental regimes. In the first, using GPT-3.5-turbo as judge across three open-ended video QA benchmarks, Oracle-RLAIF achieved the following improvements over VLM-RLAIF:"),
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
      body("The second and more rigorous experiment used Video-MME — a multiple-choice benchmark designed to prevent data leakage. Here Oracle-RLAIF achieved a +6.2% overall accuracy improvement, with particularly large gains in Temporal Perception (+21.2%), Action Recognition (+11.7%), and Object Reasoning (+11.2%). These are precisely the task types that require understanding of causally ordered events over time — exactly the domain where rank-based optimization of video responses should provide the strongest signal."),
      spacer(),
      body("The only categories where Oracle-RLAIF underperformed VLM-RLAIF were Spatial Perception (-2.6%) and Spatial Reasoning (-3.8%). The authors hypothesize that these tasks depend on implicit spatial cues that are difficult to rank reliably, and may require architectural improvements (such as stronger spatial encoders) rather than fine-tuning strategy changes."),

      h2("5.2 RL-Theoretic Interpretation of Results"),
      body("From a reinforcement learning perspective, the pattern of results is theoretically coherent. Temporal and action-based tasks have relatively clear ground-truth quality orderings — one response that correctly identifies the sequence of events is unambiguously better than one that does not. This makes Oracle ranking reliable, which in turn makes GRPOrank's advantage estimates well-calibrated."),
      spacer(),
      body("In contrast, spatial and abstract tasks involve subjective or ambiguous quality distinctions. When the Oracle's rankings are noisy or inconsistent (because the task itself is ambiguous), the rank-based training signal introduces noise into the policy gradient, potentially causing worse outcomes than a scalar reward model that can express nuanced quality differences."),
      spacer(),
      body("This finding illustrates a general principle in RL: the quality of the learning signal is directly tied to the reliability of the reward function. Oracle-RLAIF's rank-based signal is most powerful when ranking is unambiguous, and less effective when quality differences are subtle or subjective."),

      divider(),

      // ─── 6. RELATION TO BROADER RL LITERATURE ─────────────────────────────────
      h1("6. Relation to Broader RL Literature"),
      body("Oracle-RLAIF sits at the intersection of several active research threads in reinforcement learning for language models:"),
      spacer(),
      bodyMixed([{ text: "RLHF and RLAIF: ", bold: true }, { text: "The foundational framework (Ouyang et al., 2022; Bai et al., 2022; Lee et al., 2023) established preference-based RL as the standard for language model alignment. Oracle-RLAIF advances this by removing the need for a carefully trained scalar reward model." }]),
      spacer(),
      bodyMixed([{ text: "GRPO and DeepSeek-R1: ", bold: true }, { text: "GRPO (Shao et al., 2024) was originally developed to align mathematical reasoning models, proving that group-relative advantage estimation outperforms PPO in many LLM settings. Oracle-RLAIF extends GRPO to the rank-signal domain, adding the nDCG penalty as a principled way to encode ordinal feedback." }]),
      spacer(),
      bodyMixed([{ text: "Listwise Learning to Rank: ", bold: true }, { text: "The use of nDCG as an optimization target is well-established in information retrieval (learning-to-rank literature). Oracle-RLAIF is, to the best of the authors' knowledge, the first work to incorporate listwise ranking metrics directly into a policy gradient objective, bridging ranking theory and RL." }]),
      spacer(),
      bodyMixed([{ text: "DPO and Offline Preference Methods: ", bold: true }, { text: "Direct Preference Optimization (Rafailov et al., 2023) and related methods (RRHF, RAFT) optimize for preferences without an explicit reward model. However, these methods operate in a supervised or offline RL regime. Oracle-RLAIF differs by operating fully online — generating fresh responses at each training step and updating immediately — which enables it to benefit from on-policy exploration." }]),

      divider(),

      // ─── 7. CONCLUSION ─────────────────────────────────────────────────────────
      h1("7. Conclusion"),
      body("Oracle-RLAIF represents a well-motivated and practically impactful contribution to reinforcement learning for multi-modal video language models. By replacing scalar reward modeling with Oracle-based ranking, the framework removes one of the most expensive and fragile components of existing RLAIF pipelines. By introducing GRPOrank — a rank-adapted policy gradient algorithm grounded in nDCG theory — the authors provide a principled algorithmic mechanism to translate ordinal feedback into effective policy updates."),
      spacer(),
      body("From the perspective of a reinforcement learning student, the paper offers an excellent case study in how RL principles — MDP formulation, policy gradient optimization, advantage estimation, trust region constraints, and reward shaping — translate to real-world AI alignment problems. The transition from PPO to GRPO to GRPOrank traces a clear evolution in how the field has learned to handle the specific challenges of language model fine-tuning: sparse episodic reward, the need for relative comparison, and the instability introduced by reward magnitude variation."),
      spacer(),
      body("The empirical results validate the approach convincingly on temporally and causally grounded tasks, while the identified limitations in spatial and abstract reasoning point toward productive directions for future work — including stronger spatial encoders, multi-modal Oracle models, and GRPOrank applied across audio-visual tasks."),
      spacer(),

      infoBox("Key Takeaways for RL Students",
        "1. The paper shows how RLHF/RLAIF can be formalized as a standard MDP with a sequence-level, episodic reward.\n\n2. GRPO's insight — using within-group normalization to eliminate the value function — directly enables the shift from scalar reward to ordinal rank.\n\n3. nDCG is a principled way to encode position-sensitive rank quality into a continuous, bounded, and differentiable-compatible signal for policy gradient methods.\n\n4. The zero-sum property of GRPOrank ensures that the policy update always involves a meaningful trade-off between promoting good responses and suppressing bad ones.\n\n5. The paper's pattern of results — strong gains on clear-quality tasks, weaker gains on ambiguous tasks — is a direct consequence of the general RL principle that learning quality depends on reward signal reliability."),

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
