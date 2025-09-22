# Magical Girl Witch Trials Translation Project

With a script to
- Parse source localization files to XLSX
- Build localization files from XLSX
- Bundle localization files into Unity bundles
- Machine-translate the XLSX

It's recommended to use Pycharm for the pre-configured run configurations.

# Localization Guidelines for Magical Girl Visual Novel

## Project
Localization of the Magical Girl Visual Novel

## Goal
Create a complete, engaging, and faithful Vietnamese translation that preserves the tone, context, and character personalities of the original Japanese work while ensuring a natural and polished reading experience.

## Translation Guidelines

### Content Preparation
- **Tag Removal**: Before translation, remove all XML tags (e.g., ruby tags) except `<br>` and `<link>` tags, which are handled based on the processing type (standard or special). Concatenate the content seamlessly unless specified otherwise in special processing rules.
- **Line Elimination**: Completely eliminate lines starting with ">" (e.g., "> Ema: |#0101Adv02_Ema001|","> @print |#0101Adv04_Narrative111| speed:0.05 waitInput! wait!", "> ＠備考◆配信中", "> @choice |#0101Adv05_Choice001| button:ChoiceButtons/Adv/Bad", "> @toast "|#Toast_001|", "> Unknown: |#0101Adv02_Unknown001|", "> Sherry: |#0101Trial01_Sherry009|", "> Ema: |#0101Bad02_Ema006|", "> Ema: |#CommonBad01_Ema001|", "> @choice |#0101Trial01_Choice003|", "> @choice |#0101Trial01_Choice001| button:ChoiceButtons/Trial/Objection handler:Trial", "> @choice |#Common_Return| button:ChoiceButtons/Trial/Cancel handler:Trial"). These lines typically indicate speaker or metadata and should not appear in the translated output.
- **Content Retention and Cleaning**:
  - For lines not containing "> @printDebate pos:", remove all `<br>` tags during the cleaning phase and concatenate the Japanese dialogue or descriptions into a single, cohesive string to prepare for translation.
  - For lines containing "> @printDebate pos:", remove all `<br>` and `<link>` tags during the cleaning phase to create a clean text string for translation, as specified in special processing rules. Retain the original structure (positions of `<br>` and `<link>`) to restore them after translation.
- **Special Processing for Files with Format `Act01_Chapter01_Trial00`**:
  - If a line contains "> @printDebate pos:", apply the following steps:
    1. **Clean the Content**: Remove all XML tags (including `<br>` and `<link>`) and metadata (e.g., "|#0101Trial01_Sherry003|"). Retain only the main Japanese dialogue or description as a clean text string. Record the positions of `<br>` and `<link>` tags in the original line for restoration after translation.
    2. **Count `<br>` Tags**: Count the number of `<br>` tags in the original line before cleaning.
    3. **Translate the Content**: Translate the cleaned text into Vietnamese, preserving the meaning and tone.
    4. **Insert `<br>` Tags in Translation**: Based on the recorded positions and context, insert the same number of `<br>` tags into the translated text to match the original structure.
    5. **Retain `<link>` Structure**: Restore the `<link>` structure from the original line, placing it in the corresponding position in the translated text as `<link="Objection_01_01_01_01"><size=75%>[translated equivalent]</size></link><br>`.
    6. **Insert `<size=75%>` Tags**: Add a `<size=75%>` tag at the start of the translated line and immediately after each `<br>` tag.
  - **Example**:
    - Original: `"> @printDebate pos:65,35 |#0101Trial01_Sherry003| 死体の後ろ、<br>そして周囲……<br>そこに<link="Objection_01_01_01_01">血で描かれた</link><br>蝶の絵が描かれているんです！"`
    - Step 1 (Clean): Remove metadata and all tags: `死体の後ろ、そして周囲……そこに血で描かれた蝶の絵が描かれているんです！`
    - Step 2 (Count `<br>`): There are 3 `<br>` tags in the original line (after "死体の後ろ、", after "そして周囲……", and after the `<link>`).
    - Step 3 (Translate): Translate the cleaned content: "Phía sau thi thể, và xung quanh... nơi đó có hình vẽ được vẽ bằng máu, một bức tranh con bướm được vẽ!"
    - Step 4 (Insert `<br>`): Insert 3 `<br>` tags based on the original positions: After "Phía sau thi thể", after "và xung quanh...", and after the `<link>` structure.
    - Step 5 (Retain `<link>`): Restore the `<link>` structure: `<link="Objection_01_01_01_01"><size=75%>hình vẽ được vẽ bằng máu</size></link><br>`.
    - Step 6 (Add `<size=75%>`): Add `<size=75%>` tags at the start and after each `<br>` tag.
    - Final Translated Output: `<size=75%>Phía sau thi thể, <br><size=75%>và xung quanh... <br><size=75%>nơi đó có <link="Objection_01_01_01_01"><size=75%>hình vẽ được vẽ bằng máu</size></link><br><size=75%>một bức tranh con bướm được vẽ!`
- **Standard Processing** (non-special cases):
  - For lines not containing "> @printDebate pos:", remove all `<br>` tags during the cleaning phase and concatenate the content into a single string. In the final translation, ensure the output is a cohesive sentence or paragraph without `<br>` tags for natural flow in Vietnamese.
  - **Example**:
    - Original: `"> Ema: |#0101Adv02_Ema001| 何を考えていますの！？<br>わたくしをこんなところに閉じ込めるなんて！"`
    - After cleaning: `何を考えていますの！？わたくしをこんなところに閉じ込めるなんて！`
    - Translation Steps:
      - Step 1: Translate the cleaned content: "Cậu đang nghĩ gì vậy!? Nhốt mình vào một nơi như thế này ư!"
      - Step 2: Ensure no `<br>` tags in the final translation, maintaining a single cohesive sentence.
    - Final Translated Output: `Cậu đang nghĩ gì vậy!? Nhốt mình vào một nơi như thế này ư!`

### Comparison of Original and Intermediate Translation
- **Consistency Check**: Compare the Chinese translation (Translated Value) with the original Japanese (Original Value) to ensure accuracy in meaning, tone, and context.
- **Punctuation Rules**:
  - Use Vietnamese punctuation conventions (e.g., ".", ",") in the final translation.
  - Retain the Japanese dash "――" as is.
  - Replace original and Chinese ellipsis characters (e.g., "……") with the standard Vietnamese ellipsis "...".
- **Purity of Translation**: The final Vietnamese translation must be free of any Chinese or Japanese characters (e.g., "。", "、", kanji, hiragana, katakana) or words, except for the permitted "――". If any are detected, revise to ensure a pure Vietnamese translation containing only the main content, unless specified in special processing rules (e.g., `<link>` and `<size=75%>` tags).
- **Tag Restriction**: In the final translation, `<br>` tags are only permitted in lines processed under the `> @printDebate pos:` special processing rules. For all other cases, `<br>` tags must be removed, and the content must be concatenated seamlessly.
- **Personal Pronouns**: Replace Japanese pronouns (e.g., boku, watashi, atashi, uchi, wagahai, watakushi) with appropriate Vietnamese pronouns (e.g., tôi, mình, tao, ta, chị) as outlined in the **Character Pronoun Guidelines**.
- **Speaker Identification**: Lines starting with ">" (e.g., "> Ema: |#0101Adv03_Ema004|", "> Sherry: |#0101Trial01_Sherry009|", "> Ema: |#0101Bad02_Ema006|", "> Ema: |#CommonBad01_Ema001|") indicate the speaker. Track these to ensure the dialogue reflects the correct character’s voice and personality.

### Preserving Structure and Tone
- **Retention of Elements**: Preserve placeholders, variables, control codes, and line breaks (e.g., `<br>`), only as specified in special processing rules for `> @printDebate pos:` lines.
- **Tone and Honorifics**: Maintain the tone, voice, and Japanese honorifics (-kun, -san, -chan) for each character. Honorifics may be retained as is, translated (e.g., "cậu" for -kun, "chị" for -san), or omitted based on context, especially for older or superior characters.

### Quality Notes
- **Natural Flow**: Ensure the Vietnamese translation reads smoothly and naturally, as if it were a published novel. Adjust phrasing for fluency while preserving all details, avoiding overly literal translations.
- **Full Meaning**: Convey the complete meaning of the original text with sentences that are sufficiently long to avoid awkwardness and maintain the intended nuance.
- **Kanji Interpretation**: Analyze the meaning of Kanji in the original text. Use Sino-Vietnamese words (e.g., "bi thương" for 悲, meaning sorrow) when they enhance clarity or tone.
- **Final Review**: Conduct a thorough review to eliminate errors, ensuring:
  - No lines starting with ">".
  - No XML tags (except `<link>` and `<size=75%>` in `> @printDebate pos:` lines).
  - No `<br>` tags in non-`> @printDebate` lines.
  - No Chinese or Japanese characters or words (except "――") remain in the final translation.
- **Character Names and Terminology**: Strictly adhere to the provided list of names and terms without alteration or creation of new ones.

## Character Names and Terminology
- 桜羽 エマ: Sakuraba Ema (Surname: Sakuraba, Given Name: Ema)
- 蓮見 レイア: Hasumi Leia (Surname: Hasumi, Given Name: Leia)
- 氷上 メルル: Hikami Meruru (Surname: Hikami, Given Name: Meruru)
- 宝生 マーゴ: Houshou Margo (Surname: Houshou, Given Name: Margo)
- 城ケ崎 ノア: Jougasaki Noa (Surname: Jougasaki, Given Name: Noa)
- 黒部 ナノカ: Kurobe Nanoka (Surname: Kurobe, Given Name: Nanoka)
- 夏目 アンアン: Natsume An-An (Surname: Natsume, Given Name: An-An)
- 二階堂 ヒロ: Nikaidou Hiro (Surname: Nikaidou, Given Name: Hiro)
- 佐伯 ミリア: Saeki Miria (Surname: Saeki, Given Name: Miria)
- 沢渡 ココ: Sawatari Coco (Surname: Sawatari, Given Name: Coco)
- 紫藤 アリサ: Shitou Alisa (Surname: Shitou, Given Name: Alisa)
- 橘 シェリー: Tachibana Sherry (Surname: Tachibana, Given Name: Sherry)
- 遠野 ハンナ: Toono Hanna (Surname: Toono, Given Name: Hanna)
- ゴクチョー: Gokuchou
- 看守: quản ngục
- なれはて: Narehate
- スマホ: điện thoại
- 魔女: ma nữ
- 魔法: ma pháp
- 大魔女: đại ma nữ
- 魔女図鑑: bách khoa ma nữ
- 魔女因子: nhân tố ma nữ
- 魔女化: ma nữ hóa
- 処刑: tử hình
- 牢屋: nhà tù
- 火精の間: phòng hỏa tinh
- 地精の間: phòng địa tinh
- 塀: hàng rào
- 倉庫: nhà kho
- 医務室: phòng y tế
- サンルーム: phòng tắm nắng
- 物置: phòng chứa đồ
- 水精の間: phòng thủy tinh
- ゲストハウス前: trước nhà khách
- WWC: nhà vệ sinh
- 厨房: nhà bếp
- ラウンジ: phòng chờ
- シャワールーム: phòng tắm
- 中庭: sân trong
- 食堂: phòng ăn
- 裁判所: tòa án
- 湖方面: khu vực hồ
- 応接間: phòng khách
- 玄関ホール: lối đi vào
- 裁判所前通路: hành lang trước tòa án
- 花畑方面: khu vực vườn hoa
- 牢屋敷前: trước nhà tù
- 2Fホール: sảnh tầng 2
- 娯楽室: phòng giải trí
- 図書室: thư viện
- 監房: phòng giam
- 焼却炉: lò thiêu
- 懲罰房: phòng xử phạt
- 分解されたパーツ: những bộ phận bị tháo rời
- アンアンのスケッチブック: sổ phác thảo của An-An
- 血の付いたリボン: dải ruy băng dính máu
- ほうき: cây chổi
- カラースプレー: sơn xịt màu
- フルーツの写真: bức ảnh trái cây
- 配信アーカイブ: lưu trữ phát sóng
- 城ケ崎ノアの死体写真: ảnh thi thể của Jougasaki Noa
- 床の傷跡: vết nứt trên sàn
- ボウガンの矢: mũi tên của cây nỏ
- 床に描かれた蝶: con bướm được sẽ trên sàn
- 拷問ショーの案内: thông báo về chương trình tra tấn
- 

## Character Descriptions
### Sakuraba Ema (15 years old)
**Personality**: Altruistic, cheerful, clumsy, energetic, friendly, hardworking, honest, kind, optimistic, intelligent.  
**Description**: Sakuraba Ema is a friendly girl with a unique, approachable vibe. Despite her intelligence and excellent reasoning skills, Ema often deliberately makes mistakes out of fear of being disliked. Her clumsiness makes her someone others tend to look after, but behind her radiant smile lies a lonely heart yearning for warmth and connection. Ema strives to help others with persistence and selflessness, unafraid to face challenges. Her fear of isolation drives her to maintain an optimistic facade, but her sincerity and relentless effort win the hearts of those around her.

### Hasumi Leia (16 years old)
**Personality**: Cheerful, confident, friendly, hardworking, kind, mature, perceptive, protective, resilient, tomboy.  
**Description**: Hasumi Leia is a renowned actress, particularly adored by female fans for her captivating androgynous charm. With a strong tomboy style, she stands out in any crowd, but beneath her confident exterior is a warm heart always ready to protect those around her. Leia works diligently, carrying a deep sense of responsibility and maturity beyond her years. Her keen observation skills allow her to notice subtle details, and her resilience helps her overcome any challenge. Leia is not just a star on screen but also a reliable pillar for those she cares about.

### Hikami Meruru (15 years old)
**Personality**: Gentle, emotional, kind, lonely, insecure, pessimistic, sensitive, shy, stammering, reserved, timid.  
**Description**: Hikami Meruru is an extremely shy and anxious girl who quietly observes others from the shadows. With a sensitive heart, she easily picks up on others’ emotions but hides her own thoughts. Meruru is deeply insecure, often thinking negatively about herself and feeling unworthy. Her eyes always seem on the verge of tears, easily triggered by the smallest things. Despite this, her quiet kindness and concern for others are a faint light waiting for someone to notice and draw her out of her lonely shell.

### Houshou Margo (15 years old)
**Personality**: "Ara ara," cautious, curious, dishonest, suspicious, friendly, fond of female-female romance, money-loving, mysterious, relaxed, romantic, secretive.  
**Description**: Houshou Margo is a mysterious girl who appears with a friendly smile and a seductive "ara ara" tone, easily captivating those around her. However, behind her kind and romantic facade, Margo trusts no one, keeping her thoughts and motives hidden. She never takes things too seriously, maintaining a relaxed and carefree attitude in any situation. With a particular fondness for money and female-female romance, paired with an insatiable curiosity, Margo is like an elusive breeze, hiding many secrets waiting to be uncovered.

### Jougasaki Noa (15 years old)
**Personality**: Absent-minded, cheerful, eccentric, friendly, kind, moody, secretive, mysterious, refers to self in third person.  
**Description**: Jougasaki Noa is a talented street artist whose eccentric and dreamy demeanor makes her impossible to ignore. Her artworks are famous worldwide, yet no one knows Noa is their creator. With her mesmerizing ability to manipulate liquids, she delivers vibrant art performances that mirror her unpredictable personality. Though friendly and kind, Noa keeps her distance, hiding her true emotions and identity behind a carefree smile. Her mood swings make her an enigmatic figure, like an unfinished painting full of allure.

### Kurobe Nanoka (17 years old)
**Personality**: Mysterious, distant, cold, vengeful, wise, independent, reserved.  
**Description**: Kurobe Nanoka is a mysterious girl harboring deep resentment toward the prison where she is confined. Silent and aloof, she rarely speaks, choosing to push others away to conceal her inner thoughts. Nanoka seems to know more than anyone about the prison’s secret rules and structure but keeps this knowledge to herself, always acting alone. Her independence and cold gaze make her a puzzling enigma, but beneath her tough exterior may lie secrets and wounds waiting to be revealed.

### Natsume An-An (15 years old)
**Personality**: Blunt, loves cinema, moody, shy, reserved, lonely.  
**Description**: Natsume An-An is a reclusive girl who rarely speaks, often communicating through handwritten notes in her sketchbook. When writing, she uses a classical style, reflecting her unique personality and attachment to old-fashioned values. Though shy and reserved, An-An has surprising moments of bluntness, especially when discussing her passion for cinema. Her mood swings make her like a character from a film—both relatable and elusive. Her sketchbook is not just a means of communication but a window into her mysterious inner world.

### Nikaidou Hiro (15 years old)
**Personality**: Arrogant, confident, honorable, perfectionist, perceptive, serious, hot-tempered, cunning, strict, holds grudges.  
**Description**: Nikaidou Hiro embodies perfection, a girl with outstanding achievements and no apparent flaws. With the elegant and serious demeanor of a "Yamato Nadeshiko," she maintains a refined and impeccable image. However, behind her traditional beauty lies a strong will and unwavering belief in justice. Hiro is unforgiving toward what she deems "wrong," quick to show her temper and determination to correct injustices. Though confident and sometimes arrogant, her cunning and sense of honor make her irresistibly captivating. For Hiro, everything must be perfect—from herself to the world around her.

### Saeki Miria (15 years old)
**Personality**: Blunt, loves cinema, timid, melancholic, kind, insecure, old-fashioned, reserved, eccentric, shy.  
**Description**: Saeki Miria has a striking gyaru appearance with pale skin, but her shy and melancholic personality starkly contrasts her bold look. Despite her eye-catching style, Miria is quiet, easily frightened, and rarely speaks. Her soft, sorrowful tone carries a gloominess unusual for her age, as if she stepped out of an old film. Though sometimes impolite, Miria has a kind heart, quietly caring for others. Her love for cinema and old-fashioned ways make her an eccentric figure—both familiar and puzzling, like a painting with contrasting colors.

### Sawatari Coco (15 years old)
**Personality**: Fake cheerfulness, cold, cruel, selfish, arrogant, sharp-tongued, hot-tempered, cowardly, lazy, moody.  
**Description**: Sawatari Coco is a vibrant streamer with a radiant on-screen persona, captivating audiences with her fake smile and energy. Off-screen, however, Coco is cold, venomous, and harbors deep resentment toward the world. She doesn’t hesitate to use her sharp tongue to attack others, except for herself and her beloved "oshi" (idol). Her mood swings and selfish nature make her an emotional storm, both alluring and terrifying. Despite her laziness and cowardice, her overconfidence and acting talent keep people captivated by the secrets she hides.

### Shitou Alisa (15 years old)
**Personality**: Antisocial, suspicious, hot-tempered, insecure, rude, passionate.  
**Description**: Shitou Alisa is a rebellious runaway with a troubled past. Her intimidating appearance and rude demeanor make others wary, and she’s quick to fight over the slightest provocation. Alisa resents the world but hates herself most of all. Despite her tough and distrustful exterior, her passionate spirit simmers beneath, waiting for a chance to shine positively. Alisa’s contradictory nature makes her a complex puzzle—both intimidating and intriguing, with hidden pain waiting to be understood.

### Tachibana Sherry (15 years old)
**Personality**: Carefree, cheerful, confident, curious, friendly, perceptive, stubborn, amoral, eccentric.  
**Description**: Tachibana Sherry calls herself a "great detective," always appearing with a bright smile and an unstoppable free spirit. She dives into any adventure that piques her interest, without hesitation or worry. Her boundless curiosity draws her to everything unusual, from minor mysteries to complex puzzles. Though friendly and adept at recognizing familiar tropes from films or books, Sherry lacks a moral compass, acting on whims without regard for consequences. Her charming speech and stubbornness make her a colorful whirlwind, both captivating and unpredictable, like a character from an eccentric detective novel.

### Toono Hanna (15 years old)
**Personality**: Arrogant, assertive, modern tsundere, fake, perceptive, overconfident.  
**Description**: Toono Hanna carries the air of a refined lady, speaking with a haughty tone and aristocratic vocabulary. She acts as if she hails from a prestigious family, exuding unshakable confidence and arrogance. However, beneath this perfect facade lies a secret: Hanna comes from a poor background. With a modern tsundere charm, she hides her warm heart and vulnerabilities behind a lofty exterior. Her sharp observational skills let her notice details others miss, but can her noble mask conceal her true self forever?

### Gokuchou
**Personality**: Cruel, calm, mysterious.  
**Description**: Gokuchou is a mysterious owl serving as the prison’s warden. With cold, piercing eyes and a calm demeanor, it oversees everything—from trials and prisoners to the prison itself—with unrelenting cruelty. No one knows Gokuchou’s past or motives, only that this owl holds absolute power in the shadows. Every word and action exudes authority and fear, like a looming specter silently watching everything.

## Character Pronoun Guidelines
Based on personality and age (Ema, Meruru, Margo, Noa, An-An, Hiro, Miria, Coco, Alisa, Sherry, Hanna: 15 years old; Leia: 16 years old; Nanoka: 17 years old), the pronouns for each character are as follows:
- **Gokuchou (Warden)**:  
  Self-reference: Tôi (reflects its authoritative, powerful, and neutral tone as a mysterious owl).
- **Sakuraba Ema (Friendly, clumsy)**:  
  Self-reference: Mình (replaces "boku," reflecting a friendly, gentle tone).
- **Hasumi Leia (Tomboy, mature)**:  
  Self-reference: Tôi (replaces "watashi," fitting her strong, mature demeanor).
- **Hikami Meruru (Shy, stammering)**:  
  Self-reference: Mình (replaces "watashi," suitable for her timid personality).
- **Houshou Margo (Ara ara, mysterious)**:  
  Self-reference: Tôi (replaces "watashi," emphasizing her seductive, ara ara charm).
- **Jougasaki Noa (Third-person, eccentric)**:  
  Self-reference: Noa (replaces third-person self-reference, e.g., "Noa thinks...").
- **Kurobe Nanoka (Reserved, independent)**:  
  Self-reference: Tôi (replaces minimal pronoun use, fitting her cold, aloof nature).
- **Natsume An-An (Wagahai, shy)**:  
  Self-reference: Ta (replaces "wagahai," reflecting her classical, old-fashioned style, especially in writing).
- **Nikaidou Hiro (Arrogant)**:  
  Self-reference: Tôi (replaces "watashi," reflecting her serious, elegant demeanor).
- **Saeki Miria (Third-person or watashi, melancholic)**:  
  Self-reference: Mình or Miria (replaces "watashi" or third-person, depending on mood).
- **Sawatari Coco (Sharp-tongued)**:  
  Self-reference: Tui (replaces "atashi," reflecting her confident, sharp persona).
- **Shitou Alisa (Rude)**:  
  Self-reference: Tao, occasionally Tôi (replaces "uchi," reflecting her rebellious nature, with "Tôi" in more serious contexts).
- **Tachibana Sherry (Carefree)**:  
  Self-reference: Tớ, occasionally Tôi (replaces "watashi," with "Tớ" for friendliness and a cute tone, paired with expressions like "nhé" or "nha" to reflect her playful style; "Tôi" in more formal moments).
- **Toono Hanna (Haughty)**:  
  Self-reference: Tôi (replaces "watakushi," reflecting her aristocratic demeanor).