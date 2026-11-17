# Fewer Bits, Same Brain: Quantizing LLMs

Raaj Banga

An introduction to **post-training quantization** (storing a trained model's weights with fewer bits so it fits and runs on smaller hardware), with a hands-on notebook that runs one real model, **Qwen3-4B-Instruct-2507**, at six compression levels on a laptop and measures where it breaks.

| File | What it is |
|---|---|
| `presentation.pptx` | Slide deck (18 slides, about 15 min plus a 5–7 min live demo) |
| `quantization_demo.ipynb` | The demo notebook, saved with outputs from a full run |
| `requirements.txt` | Python packages for the notebook, pinned to the exact versions used |
| `requirements-lock.txt` | Full `pip freeze` of the environment that produced `results/` |
| `tools/` | The scripts used to generate the slides (pptxgenjs) and the picture assets, kept for reference |
| `results/` | Saved numbers (`results.csv`, `weight_error.csv`), full model replies (`transcripts.json`) and charts (`figures/`) |

## What the notebook does

1. **Rounding in miniature**: eight numbers rounded to 8, 4 and 2 bits, shown on a number line.
2. **Rounding a picture**: the same idea on an image, which explains why image generators tolerate it.
3. **Outliers**: one huge value ruins rounding for everything else; per-block scales fix it.
4. **Real weights**: opens a weight matrix inside the actual model files and measures the rounding error (`E = W − Ŵ`).
5. **The real test**: loads each of the six files and measures memory, load time, speed (tokens/s) and perplexity (one paragraph plus WikiText-2), then grades two difficulty levels: easy (20 math word problems, 10 short JSON function calls) and hard (50 GSM8K problems, 8 long JSON calls with 5–6 arguments each).
6. **Charts and limitations**, plus a ~20-second **live cell** that asks the 4-bit and 1.5-bit models the same question.

## Setup (macOS, Apple Silicon)

Tested on an M4 Pro MacBook with 24 GB of memory and Python 3.14. You need about 21 GB of free disk space for the model files.

```bash
python3 -m venv .venv
source .venv/bin/activate
CMAKE_ARGS="-DGGML_METAL=on" pip install -r requirements.txt   # pinned versions
python -m ipykernel install --user --name quant-demo --display-name "Quantization Demo (.venv)"
```

To reproduce the exact environment (every sub-package), use `pip install -r requirements-lock.txt` instead (same `CMAKE_ARGS`).

`llama-cpp-python` compiles llama.cpp from source, which takes a few minutes. `CMAKE_ARGS` turns on the Apple GPU (Metal). On Linux with an NVIDIA GPU, use `CMAKE_ARGS="-DGGML_CUDA=on"` instead.

## Running it

**VS Code (Jupyter extension):** open the folder, open `quantization_demo.ipynb`, click **Select Kernel** (top right) and pick **Quantization Demo (.venv)**. `.vscode/settings.json` already points VS Code at `.venv`. Everything runs locally on your Mac.

**Terminal:**

```bash
jupyter notebook quantization_demo.ipynb
```

- **Full run**: set `RUN_MODELS = True` in the setup cell and run all cells. The first run downloads the six GGUF files (about 19.7 GB) into `models/` and then takes roughly 75 minutes on an M4 Pro (the GSM8K and WikiText-2 files are small and download on the first run too).
- **Re-plot only** (the default): leave `RUN_MODELS = False`. The notebook reads the saved `results/` instead of running the models. Parts 1–3 never need the model files.
- **Live cell**: the last cell runs on its own (no need to run the cells above) and needs only the Q4_K_M and IQ1_S files; it downloads them if they are missing.

## Model and data sources

| Item | Source |
|---|---|
| Base model | [Qwen/Qwen3-4B-Instruct-2507](https://huggingface.co/Qwen/Qwen3-4B-Instruct-2507) (Apache-2.0) |
| Quantized files | [unsloth/Qwen3-4B-Instruct-2507-GGUF](https://huggingface.co/unsloth/Qwen3-4B-Instruct-2507-GGUF): `F16`, `Q8_0`, `Q4_K_M`, `Q3_K_M`, `Q2_K`, `UD-IQ1_S` |
| Runtime | [llama.cpp](https://github.com/ggml-org/llama.cpp) through [llama-cpp-python](https://github.com/abetlen/llama-cpp-python) |
| Easy math | 20 original GSM8K-style word problems, written for this project and listed in the notebook |
| Hard math | 50 problems sampled (`random_state=0`) from the [GSM8K](https://huggingface.co/datasets/openai/gsm8k) test set (MIT) |
| JSON tasks | 10 short + 8 long original function-call requests, listed in the notebook |
| Perplexity text | One original paragraph, plus 8 × 512-token windows of the [WikiText-2](https://huggingface.co/datasets/Salesforce/wikitext) test set (CC BY-SA) |
| Image | Generated with numpy inside the notebook (no download) |

## Limitations

- The test sets are small (20 + 50 math, 10 + 8 JSON). They show where the model breaks, not 1–2% differences.
- WikiText-2 perplexity uses 8 windows of 512 tokens, not the full test set that papers report, so compare levels with each other, not with published numbers.
- One model and one family of formats (llama.cpp K-quants and I-quants). Larger models generally tolerate low bits better.
- Speeds are from Apple unified memory. The pattern carries over to NVIDIA GPUs; the exact numbers don't.

## References

The sources cited in the presentation.

- Apple. (2025). Apple Intelligence foundation language models: Tech report 2025. arXiv. https://doi.org/10.48550/arXiv.2507.13575
- Cobbe, K., et al. (2021). Training verifiers to solve math word problems. arXiv. https://doi.org/10.48550/arXiv.2110.14168
- DeepSeek-AI. (2024). DeepSeek-V3 technical report. arXiv. https://doi.org/10.48550/arXiv.2412.19437
- Dettmers, T., et al. (2022). LLM.int8(): 8-bit matrix multiplication for transformers at scale. Advances in Neural Information Processing Systems 35 (NeurIPS 2022). https://doi.org/10.48550/arXiv.2208.07339
- Dettmers, T., et al. (2023). QLoRA: Efficient finetuning of quantized LLMs. Advances in Neural Information Processing Systems 36 (NeurIPS 2023). https://doi.org/10.48550/arXiv.2305.14314
- Element Labs. (2025). LM Studio [Software]. https://lmstudio.ai
- Frantar, E., et al. (2023). GPTQ: Accurate post-training quantization for generative pre-trained transformers. Proceedings of the 11th International Conference on Learning Representations (ICLR 2023). https://doi.org/10.48550/arXiv.2210.17323
- ggml-org. (2025). llama.cpp and the GGUF file format [Software]. https://github.com/ggml-org/llama.cpp
- Li, M., et al. (2025). SVDQuant: Absorbing outliers by low-rank components for 4-bit diffusion models. Proceedings of the 13th International Conference on Learning Representations (ICLR 2025). https://doi.org/10.48550/arXiv.2411.05007
- Lin, J., et al. (2024). AWQ: Activation-aware weight quantization for on-device LLM compression and acceleration. Proceedings of Machine Learning and Systems 6 (MLSys 2024). https://doi.org/10.48550/arXiv.2306.00978
- Ma, S., et al. (2024). The era of 1-bit LLMs: All large language models are in 1.58 bits. arXiv. https://doi.org/10.48550/arXiv.2402.17764
- Merity, S., et al. (2017). Pointer sentinel mixture models (WikiText-2 dataset). Proceedings of the 5th International Conference on Learning Representations (ICLR 2017). https://doi.org/10.48550/arXiv.1609.07843
- Meta AI. (2024). The Llama 3 herd of models. arXiv. https://doi.org/10.48550/arXiv.2407.21783
- Ollama. (2025). Ollama [Software]. https://ollama.com
- OpenAI. (2025). gpt-oss-120b & gpt-oss-20b model card. arXiv. https://doi.org/10.48550/arXiv.2508.10925
- Qwen Team. (2025). Qwen3 technical report. arXiv. https://doi.org/10.48550/arXiv.2505.09388
- Rouhani, B. D., et al. (2023). Microscaling data formats for deep learning. arXiv. https://doi.org/10.48550/arXiv.2310.10537
- Xiao, G., et al. (2023). SmoothQuant: Accurate and efficient post-training quantization for large language models. Proceedings of the 40th International Conference on Machine Learning (ICML 2023), PMLR 202. https://doi.org/10.48550/arXiv.2211.10438
