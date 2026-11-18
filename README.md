# intelligence-studio

## أدوات Claude Code الخمسة

| # | الأداة | وش تسوي | كيف مركّبة هنا |
|---|---|---|---|
| 1 | **OmniRoute** | بوابة وحدة لأكثر من 50 مزوّد AI، وتحوّلك تلقائياً لمزوّد ثاني إذا خلص الحد | `npm i -g omniroute` (عبر `scripts/install-tools.sh`) |
| 2 | **claude-mem** | ذاكرة دائمة: يتذكر وش سويت في الجلسات السابقة | Plugin في `.claude/settings.json` |
| 3 | **Headroom** | يضغط اللوقات والملفات قبل ما توصل لـ Claude ويقلل التوكنز | `pip install "headroom-ai[code]"` |
| 4 | **claude-code-setup** | Plugin رسمي من Anthropic: يحلل مشروعك ويقترح hooks وskills وMCP | Plugin في `.claude/settings.json` |
| 5 | **Task Observer** | يلاحظ الأغلاط المتكررة ويحفظها كقواعد في SKILL.md | Skill في `.claude/skills/task-observer/` |

### التثبيت

```bash
./scripts/install-tools.sh
```

لما تفتح Claude Code داخل هذا المشروع، بيطلب منك الموافقة على الـ marketplaces والـ plugins المذكورة في `.claude/settings.json`.

### الاستخدام

- **claude-code-setup**: اكتب داخل Claude Code: `recommend automations for this project`
- **claude-mem**: يشتغل تلقائياً (hooks). يحتاج Node 18+.
- **Headroom**: `headroom wrap claude` (يشغّل Claude Code من خلال proxy الضغط)
- **OmniRoute**: `OMNIROUTE_SERVER_HOST=127.0.0.1 omniroute serve`، ثم افتح `http://localhost:20128` وأضف المزوّدين، ثم `omniroute setup` لربطه بـ Claude Code.
- **Task Observer**: يشتغل أثناء المهام. للأفضل أضف تعليمات تفعيله في `CLAUDE.md` (شوف `references/environments.md` داخل الـ skill).

### تنبيهات

- **OmniRoute** يشتغل افتراضياً على `0.0.0.0` بدون مفتاح API. شغّله على `127.0.0.1` أو فعّل `REQUIRE_API_KEY=true`. كمان فيه مقالات تنتقد أمانه، وبعض المزوّدين "المجانيين" ممكن يخالفون شروط الاستخدام. راجعه قبل ما تحط مفاتيح حقيقية.
- أرقام مثل "1.6B توكن ببلاش" و"120K → 18K" ادعاءات تسويقية وما تأكدت منها.
