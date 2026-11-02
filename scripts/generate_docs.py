import os
import zipfile
from xml.sax.saxutils import escape

BASE_DIR = r"C:\Users\user\.gemini\antigravity\scratch\ai-dnd-trpg-companion"
DOCX_PATH = os.path.join(BASE_DIR, "report", "project_report.docx")
PPTX_PATH = os.path.join(BASE_DIR, "slides", "presentation_1102.pptx")


def build_docx(output_path):
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    paragraphs = [
        ("title", "AI 雙視窗《龍與地下城》(D&D) 跑團與獨立隊友系統 — 專案報告"),
        ("subtitle", "日期：2026 年 11 月 2 日 ｜ 主題：基於多代理架構（Multi-Agent）之桌上角色扮演遊戲（TRPG）協作系統"),
        ("h1", "一、 做什麼（What We Built）"),
        ("p", "本專案打造了一個「主聊天室擔任地下城主（DM）＋ 獨立子代理視窗擔任 AI 隊友」的《龍與地下城》（D&D 5e）文字跑團系統。"),
        ("p", "在傳統單一聊天室與 AI 進行 TRPG 時，若讓同一個 AI 同時扮演 DM 與隊友，往往會遇到兩大問題：（1）上帝視角劇透（Metagaming）：AI 隊友因為與 DM 共用記憶，會提前知道前面轉角有陷阱或埋伏幾隻哥布林；（2）敘事角色混淆：同一個視窗同時輸出規則判定與隊友吐槽，缺乏真實夥伴的互動感。"),
        ("bullet", "主聊天室（地下城主 DM）：負責《失落的礦坑》世界觀描寫、環境氛圍營造、怪物行動，以及維護「隊伍狀態欄」（HP、AC、法術位）與進行 D20 檢定。"),
        ("bullet", "獨立子視窗（AI 隊友：高等精靈女法師「萊拉・星語 Lyra」）：透過 Subagent 技術啟動另一個獨立對話空間。萊拉僅接收 DM 傳遞的「當前玩家可見場景」，並根據自己傲嬌但可靠的個性與法師技能（如《魔法伎倆》、《法師護甲》、《火焰箭》）做出真實反應，再由主視窗同步轉播給人類玩家。"),
        ("h1", "二、 給誰用（Target Audience）"),
        ("bullet", "1. 找不到固定團友的「孤狼玩家」：傳統 D&D 需要湊齊 4～5 人並耗時數小時；本系統讓單人玩家隨時打開聊天室，就有專屬 DM 加上一位有血有肉、會吐槽也會保護你的 AI 隊友陪伴冒險。"),
        ("bullet", "2. 想嘗試 TRPG 但害怕厚重規則書的新手：系統將繁瑣的創角流程轉化為直覺的互動選單，並在頂部自動維護狀態欄，讓新手專注於故事決策與扮演。"),
        ("bullet", "3. Multi-Agent 遊戲應用開發者：展示如何利用「主從式代理（Orchestrator-Subagent）」實現資訊不對稱（Information Asymmetry）的互動體驗，可延伸至狼人殺、劇本殺（LARP）等遊戲。"),
        ("h1", "三、 怎麼跟 AI 一起做（How We Collaborated with AI）"),
        ("bullet", "步驟一（提出核心需求）：向 AI 提出「我想在這個聊天室玩龍與地下城你能幫我嗎」，建立基礎跑團引導框架。"),
        ("bullet", "步驟二（腦力激盪雙視窗隊友機制）：進一步詢問「能不能開另一個聊天室當我的隊友？」，AI 規劃出在後台啟動獨立 Subagent 並自動轉播台詞的「方案 A」。"),
        ("bullet", "步驟三（結構化選單快速配置）：透過互動式選單（ask_question），快速選定「人類戰士/聖騎士（玩家，HP 12, AC 18）＋ 傲嬌精靈女法師萊拉（AI 隊友，HP 7, AC 15）＋ 經典劇本《失落的礦坑》」。"),
        ("bullet", "步驟四（動態定義與召喚子代理）：AI 自動呼叫 define_subagent 建立 dnd_companion_lyra 角色模板（嚴格限制玩家視角、禁止自行編造劇情結果），並透過 invoke_subagent 傳送第一章「三豬小徑死馬伏擊」場景，獲得萊拉生動且符合 D&D 規則的開場演出："),
        ("quote", "【獨立視窗「萊拉」實際生成的開場反應】：「喂！停下馬車，別再往前傻衝了！你看那兩匹馬……那種粗糙噁心的黑羽箭，絕對是哥布林的傑作！……本小姐的防護（施放一環法術《法師護甲》，AC 提升至 15）已經準備好了！你這個穿鐵罐頭的肌肉腦袋打算怎麼辦？要是敢讓那些髒兮兮的小怪物靠近我半步，你就死定了！」"),
        ("h1", "四、 查了什麼（What We Researched & Referenced）"),
        ("bullet", "1. D&D 5e 系統參考文件（SRD 5.1 規則）：查證 1 級聖騎士（鎖子甲＋盾牌 AC 18、聖療術 5 點）與 1 級高等精靈法師的數值計算，包含戲法《魔法伎倆》（清潔法袍泥巴，不耗法術位）、《火焰箭》，以及一環法術《法師護甲》（13 + 敏捷調整值 2 = AC 15，消耗 1 格法術位）。"),
        ("bullet", "2. D&D 官方經典入門戰役《失落的礦坑》（Lost Mine of Phandelver）：參考第一章〈哥布林之箭〉（Goblin Arrows）開場——矮人僱主岡德林・尋石者（Gundren Rockseeker）與武士希爾達（Sildar Hallwinter）先行出發，玩家於三豬小徑（Triboar Trail）遭遇插滿黑羽箭的兩匹死馬。"),
        ("bullet", "3. 多代理系統之記憶隔離（Context Isolation）：驗證主代理與子代理之間的通訊邊界，確保 invoke_subagent 僅傳遞過濾後的玩家感官資訊，防止 AI 隊友偷看 DM 的隱藏劇本。"),
        ("h1", "五、 下一步（Next Steps）"),
        ("bullet", "1. 推進三豬小徑遭遇戰與先攻系統（Initiative Tracker）：實作正式的 D20 感知檢定、隱匿對抗與回合制戰鬥，讓玩家與萊拉共同擊退哥布林伏兵。"),
        ("bullet", "2. 擴充為「多視窗四人標準小隊」：除了精靈法師萊拉外，再召喚「豪爽的矮人狂戰士巴爾」與「嘴砲提夫林盜賊芬恩」兩個獨立 Subagent 視窗，讓多個 AI 隊友之間也能互相鬥嘴與打配合。"),
        ("bullet", "3. 整合視覺化戰場與自動存檔（JSON State Persistence）：將角色卡、物品欄、法術位消耗與劇情進度自動存成 JSON 狀態檔，並結合互動式網頁骰子與地圖面板。")
    ]

    body_xml = []
    for style, text in paragraphs:
        esc = escape(text)
        if style == "title":
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:b/><w:color w:val="1E1B4B"/><w:sz w:val="36"/></w:rPr>
                <w:t>{esc}</w:t>
              </w:r>
            </w:p>''')
        elif style == "subtitle":
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:spacing w:after="240"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:color w:val="4F46E5"/><w:sz w:val="20"/></w:rPr>
                <w:t>{esc}</w:t>
              </w:r>
            </w:p>''')
        elif style == "h1":
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:b/><w:color w:val="312E81"/><w:sz w:val="28"/></w:rPr>
                <w:t>{esc}</w:t>
              </w:r>
            </w:p>''')
        elif style == "bullet":
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:ind w:left="360"/><w:spacing w:after="100"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:sz w:val="22"/></w:rPr>
                <w:t>• {esc}</w:t>
              </w:r>
            </w:p>''')
        elif style == "quote":
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:ind w:left="480" w:right="240"/><w:spacing w:before="100" w:after="140"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:i/><w:color w:val="92400E"/><w:sz w:val="21"/></w:rPr>
                <w:t>{esc}</w:t>
              </w:r>
            </w:p>''')
        else:
            body_xml.append(f'''
            <w:p>
              <w:pPr><w:spacing w:after="120"/></w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Microsoft JhengHei" w:eastAsia="Microsoft JhengHei"/><w:sz w:val="22"/></w:rPr>
                <w:t>{esc}</w:t>
              </w:r>
            </w:p>''')

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {''.join(body_xml)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)


def make_slide_xml(title, bullets, footer="2026/11/02 點子分享 — AI 雙視窗 D&D 跑團系統"):
    bullet_paras = []
    for b in bullets:
        bullet_paras.append(f"""
        <a:p>
          <a:pPr marL="342900" indent="-342900"><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:pPr>
          <a:r>
            <a:rPr lang="zh-TW" sz="2200" dirty="0">
              <a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill>
              <a:latin typeface="Microsoft JhengHei"/>
              <a:ea typeface="Microsoft JhengHei"/>
            </a:rPr>
            <a:t>▸ {escape(b)}</a:t>
          </a:r>
        </a:p>""")

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="457200"/><a:ext cx="10820400" cy="1143000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="zh-TW" sz="3400" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="A5B4FC"/></a:solidFill>
                <a:latin typeface="Microsoft JhengHei"/>
                <a:ea typeface="Microsoft JhengHei"/>
              </a:rPr>
              <a:t>{escape(title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="1714500"/><a:ext cx="10820400" cy="4343400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/>
          <a:lstStyle/>
          {''.join(bullet_paras)}
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="6172200"/><a:ext cx="10820400" cy="457200"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="zh-TW" sz="1400" dirty="0">
                <a:solidFill><a:srgbClr val="64748B"/></a:solidFill>
                <a:latin typeface="Microsoft JhengHei"/>
                <a:ea typeface="Microsoft JhengHei"/>
              </a:rPr>
              <a:t>{escape(footer)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def build_pptx(output_path):
    slides_data = [
        (
            "🎲 AI 雙視窗《龍與地下城》(D&D) 跑團與獨立隊友系統",
            [
                "2026 / 11 / 02 點子分享簡報",
                "核心突破：用「主聊天室 (DM) ＋ 獨立子視窗 (AI 隊友)」解決單人跑團痛點",
                "結合 D&D 5e 標準規則、無劇透記憶隔離（Context Isolation）與即時角色扮演",
                "玩家扮演：人類聖騎士 ／ 獨立 AI 隊友：傲嬌高等精靈女法師【萊拉 (Lyra)】"
            ]
        ),
        (
            "1. 做什麼：為什麼需要「兩個聊天室」玩 D&D？",
            [
                "傳統單一聊天室的痛點 1（上帝視角劇透 Metagaming）：AI 隊友與 DM 共用大腦，提前知道哪裡有哥布林埋伏",
                "傳統單一聊天室的痛點 2（角色混淆）：同一視窗同時報數據又自己演隊友，缺乏真實夥伴互動感",
                "我們的方案：主聊天室擔任「地下城主 (DM)」，掌控《失落的礦坑》世界觀、怪物與 D20 檢定",
                "獨立子視窗擔任「AI 隊友萊拉 (Subagent)」，只接收玩家可見的感官描述，再將台詞同步回主視窗！"
            ]
        ),
        (
            "2. 給誰用：目標受眾與應用場景",
            [
                "🐺 找不到固定團友的「孤狼玩家」：不用苦等湊齊 4～5 人，隨時打開就有 DM 與專屬 AI 隊友陪你冒險",
                "🌱 想入坑的 TRPG 新手：透過直覺選單 30 秒創角，頂部自動維護 HP、AC 與法術位狀態欄",
                "🤖 多代理（Multi-Agent）遊戲開發者：展示「資訊不對稱（Information Asymmetry）」的記憶隔離架構",
                "可進一步延伸至：劇本殺（LARP）、狼人殺、戰術小隊文字冒險遊戲"
            ]
        ),
        (
            "3. 怎麼跟 AI 一起做：從一句話到雙代理協作",
            [
                "Step 1 提出核心需求：「我想在這個聊天室玩龍與地下城你能幫我嗎」建立跑團框架",
                "Step 2 激發架構點子：「能不能開另一個聊天室當我的隊友？」👉 選擇方案 A（召喚獨立 Subagent）",
                "Step 3 結構化選單配置：快速拍板「人類聖騎士 ＋ 傲嬌精靈女法師萊拉 ＋ 失落的礦坑劇本」",
                "Step 4 動態定義與召喚：AI 自動呼叫 define_subagent 與 invoke_subagent 啟動萊拉的獨立視窗"
            ]
        ),
        (
            "✨ 實機跑團展示：三豬小徑的伏擊與隊友萊拉",
            [
                "隊伍狀態欄：🛡️ 人類聖騎士 (HP 12/12, AC 18) ｜ 🪄 萊拉 (高等精靈法師 HP 7/7, AC 15 已開法師護甲)",
                "DM 場景：三豬小徑前方 50 呎橫躺兩匹插滿粗糙黑羽箭的死馬，兩旁灌木叢安靜得詭異",
                "萊拉獨立視窗反應：「唔……泥巴都濺到我從銀月城帶來的星紋法袍下擺了！」（施展《魔法伎倆》清污漬）",
                "萊拉進入戒備：「喂！停下馬車！那種粗糙噁心的黑羽箭絕對是哥布林的傑作！本小姐的《法師護甲》準備好了，你這個穿鐵罐頭的肌肉腦袋打算怎麼辦？」"
            ]
        ),
        (
            "4. 查了什麼：背後的規則與技術查證",
            [
                "📘 D&D 5e SRD 規則：查證 1 級聖騎士（鎖子甲+盾牌 AC 18）與 1 級精靈法師（《法師護甲》13+敏捷2 = AC 15）數值",
                "🗺️ 官方經典模組《失落的礦坑》(Lost Mine of Phandelver)：還原矮人僱主岡德林、武士希爾達與三豬小徑伏擊開場",
                "🔒 Subagent 記憶隔離設計：在萊拉的 System Prompt 中嚴格規定「只知道 DM 傳遞的場景，禁止替怪物做決定或劇透」"
            ]
        ),
        (
            "5. 下一步：未來擴充藍圖 (Next Steps)",
            [
                "🎲 1. 完整回合制戰鬥：推進三豬小徑遭遇戰，加入自動 D20 感知檢定、先攻順序（Initiative）與哥布林戰鬥",
                "🍻 2. 擴充為「多視窗四人標準小隊」：再召喚「矮人狂戰士巴爾」與「提夫林盜賊芬恩」，讓 AI 隊友之間也能鬥嘴配合",
                "💾 3. 視覺化戰場與自動存檔：將角色卡、HP、法術位與戰利品寫入 JSON 狀態檔，並整合網頁互動骰子盤與地圖"
            ]
        )
    ]

    num_slides = len(slides_data)

    # [Content_Types].xml
    ct_overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    ]
    for i in range(1, num_slides + 1):
        ct_overrides.append(
            f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(ct_overrides)}
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

    sld_id_list = []
    pres_rels_list = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
    ]
    for i in range(1, num_slides + 1):
        rid = f"rId{i+2}"
        sld_id = 255 + i
        sld_id_list.append(f'<p:sldId id="{sld_id}" r:id="{rid}"/>')
        pres_rels_list.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        )

    presentation_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {''.join(sld_id_list)}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

    presentation_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {''.join(pres_rels_list)}
</Relationships>"""

    slide_master_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"""

    slide_master_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

    slide_layout_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

    slide_layout_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

    theme_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Microsoft JhengHei"/><a:ea typeface="Microsoft JhengHei"/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft JhengHei"/><a:ea typeface="Microsoft JhengHei"/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

    slide_rel_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("ppt/presentation.xml", presentation_xml)
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels)
        zf.writestr("ppt/theme/theme1.xml", theme_xml)
        for idx, (title, bullets) in enumerate(slides_data, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", make_slide_xml(title, bullets, f"2026/11/02 點子分享 — AI 雙視窗 D&D 跑團系統 ({idx}/{num_slides})"))
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rel_xml)


if __name__ == "__main__":
    build_docx(DOCX_PATH)
    build_pptx(PPTX_PATH)
    print("Successfully generated DOCX and PPTX!")
