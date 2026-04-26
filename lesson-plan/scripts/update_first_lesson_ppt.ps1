$ErrorActionPreference = "Stop"

$source = "E:\2024-2025课程\实验动画创作\中国文化实验动画创作课件\中国文化实验创作课件\中国文化实验动画创作01.pptx"
$out = "D:\codex\lesson plan\中国文化实验动画创作25-26（2）第一课次课件.pptx"

Copy-Item -LiteralPath $source -Destination $out -Force

$msoTrue = -1
$msoFalse = 0
$msoTextOrientationHorizontal = 1

function Get-ShapeById($slide, [int]$id) {
    foreach ($shape in $slide.Shapes) {
        if ($shape.Id -eq $id) { return $shape }
    }
    return $null
}

function Set-TextById($slide, [int]$id, [string]$text, [double]$fontSize = 0) {
    $shape = Get-ShapeById $slide $id
    if ($null -ne $shape -and $shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $shape.TextFrame.TextRange.Text = $text
        if ($fontSize -gt 0) {
            $shape.TextFrame.TextRange.Font.Size = $fontSize
        }
        $shape.TextFrame.TextRange.Font.Name = "微软雅黑"
    }
}

function Add-TextBox($slide, [double]$left, [double]$top, [double]$width, [double]$height, [string]$text, [double]$fontSize = 18, [string]$fontName = "微软雅黑") {
    $box = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.TextRange.Font.Name = $fontName
    $box.TextFrame.TextRange.Font.Size = $fontSize
    $box.TextFrame.WordWrap = $msoTrue
    $box.TextFrame.AutoSize = 0
    return $box
}

function Clear-TextExcept($slide, [int[]]$keepIds) {
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $slide.Shapes.Item($i)
        if (($shape.HasTextFrame -and $shape.TextFrame.HasText) -and ($keepIds -notcontains $shape.Id)) {
            $shape.Delete()
        }
    }
}

function Set-Title($slide, [string]$title) {
    $first = $null
    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            $first = $shape
            break
        }
    }
    if ($null -ne $first) {
        $first.TextFrame.TextRange.Text = $title
        $first.TextFrame.TextRange.Font.Name = "微软雅黑"
        $first.TextFrame.TextRange.Font.Size = 23.05
    }
}

function Add-Bullets($slide, [double]$left, [double]$top, [double]$width, [double]$height, [string[]]$items, [double]$fontSize = 18) {
    $text = ($items | ForEach-Object { "• " + $_ }) -join "`r"
    return Add-TextBox $slide $left $top $width $height $text $fontSize
}

function Add-CaseLine($slide, [string]$text) {
    $box = Add-TextBox $slide 36 342 650 44 $text 14.5
    $box.TextFrame.TextRange.Font.Bold = $msoTrue
    return $box
}

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$pres = $app.Presentations.Open($out, $msoFalse, $msoFalse, $msoFalse)

# Slide 3: assessment table
$s = $pres.Slides.Item(3)
Set-Title $s "课程基本信息介绍"
for ($i = $s.Shapes.Count; $i -ge 1; $i--) {
    $shape = $s.Shapes.Item($i)
    if ($shape.Type -eq 19) { $shape.Delete() }
}
Add-TextBox $s 30 82 660 34 "考核形式：考试（作品汇报）" 18 | Out-Null
$tableShape = $s.Shapes.AddTable(3, 7, 30, 130, 664, 190)
$table = $tableShape.Table
$row1 = @("平时成绩（30%）", "平时成绩（30%）", "期中考核（30%）", "期中考核（30%）", "期中考核（30%）", "期末考试（40%）", "总成绩")
$row2 = @("课堂表现", "课堂练习", "作业1", "作业2", "实践报告", "作品汇报", "总成绩")
$row3 = @("30分", "70分", "30分", "50分", "20分", "100分", "100分")
for ($c = 1; $c -le 7; $c++) {
    $table.Cell(1, $c).Shape.TextFrame.TextRange.Text = $row1[$c-1]
    $table.Cell(2, $c).Shape.TextFrame.TextRange.Text = $row2[$c-1]
    $table.Cell(3, $c).Shape.TextFrame.TextRange.Text = $row3[$c-1]
    for ($r = 1; $r -le 3; $r++) {
        $tr = $table.Cell($r, $c).Shape.TextFrame.TextRange
        $tr.Font.Name = "微软雅黑"
        $tr.Font.Size = 13.5
        $tr.ParagraphFormat.Alignment = 2
        $table.Cell($r, $c).Shape.TextFrame.VerticalAnchor = 3
    }
}
Add-TextBox $s 45 326 630 42 "说明：平时成绩由课堂表现与课堂练习构成；期中考核由作业1、作业2与实践报告构成；期末以作品汇报进行综合评价。" 12.5 | Out-Null

# Slide 4: basic information
$s = $pres.Slides.Item(4)
Set-TextById $s 3 "课程基本信息介绍"
Set-TextById $s 12 "课程名称与开课信息"
Set-TextById $s 13 "《中国文化实验动画创作》面向动画专业开设，2025-2026学年第二学期授课，围绕中国文化、实验动画方法与微动画创作实践展开。" 14.5
Set-TextById $s 14 "学时与学分"
Set-TextById $s 15 "64学时（理论16 / 实践48）`r4学分" 14.5
Set-TextById $s 16 "学情分析情况"
Set-TextById $s 17 "学生已有动画运动规律基础，但对实验动画的材料语言、分类框架和中国文化转译方法仍需系统建立。" 14.5
Set-TextById $s 18 "课程类型"
Set-TextById $s 19 "专业基础必修课" 16

# Slide 5: course modules
$s = $pres.Slides.Item(5)
Set-Title $s "课程基本信息介绍"
Clear-TextExcept $s @(2)
Add-TextBox $s 60 80 620 32 "课程模块与教学安排" 21 | Out-Null
Add-Bullets $s 66 128 600 170 @(
    "模块一：文化实验动画概述，理论8学时 + 实践4学时",
    "模块二：中国文化实验动画作品赏析，理论8学时 + 实践8学时",
    "模块三：中期制作与资产管理，实践20学时",
    "模块四：后期制作与衍生开发，实践16学时"
) 16 | Out-Null
Add-TextBox $s 66 315 600 42 "本课次对应模块一：实验动画的定义、特征与分类导入，并衔接《材料语言测试》实践。" 15 | Out-Null

# Slide 6: objectives
$s = $pres.Slides.Item(6)
Set-TextById $s 2 "教学目标与重难点"
Set-TextById $s 15 "知识目标"
Set-TextById $s 16 "能准确说出实验动画定义，概括其区别于主流动画的核心特征，掌握主题、创作方式、材料三类观察维度。" 11.5
Set-TextById $s 17 "能力目标"
Set-TextById $s 18 "能够从形式、材料和表达方式切入分析案例，并围绕材料语言测试形成初步构思。" 11.5
Set-TextById $s 19 "素养目标"
Set-TextById $s 20 "理解中国动画学派的实验传统，建立守正创新、主动试错和文化自信意识。" 11.5
Set-TextById $s 21 "教学重点"
Set-TextById $s 22 "重点：定义、特征与三种分类维度。难点：分类交叉关系及材料语言如何参与动画表达。" 11.5

# Slide 7: contents
$s = $pres.Slides.Item(7)
Set-TextById $s 3 "目录"
Set-TextById $s 6 "课程导入与问题激活"
Set-TextById $s 9 "实验动画定义与特征"
Set-TextById $s 12 "分类框架与案例辨析"
Set-TextById $s 15 "材料语言测试任务"

# Slide 8: formal teaching section
$s = $pres.Slides.Item(8)
Clear-TextExcept $s @()
Add-TextBox $s 76 120 90 80 "02" 42 | Out-Null
Add-TextBox $s 180 124 450 78 "课堂正式教学内容" 32 | Out-Null
Add-TextBox $s 182 205 450 42 "从导入问题进入实验动画的观察框架" 17 | Out-Null

# Slides 9-11: intro comparison
$s = $pres.Slides.Item(9)
Set-Title $s "导入问题：哪一部更接近实验动画？"
Clear-TextExcept $s @(3)
Add-TextBox $s 38 78 620 40 "比较观看：商业叙事动画、视觉诗性动画与中国水墨动画在观看体验上的差异。" 18 | Out-Null
Add-CaseLine $s "影片案例：《大鱼海棠》《小蝌蚪找妈妈》；可补充：一部视觉诗歌类短片片段。"

$s = $pres.Slides.Item(10)
Set-Title $s "案例对比：《大鱼海棠》"
Clear-TextExcept $s @(3)
Add-Bullets $s 42 80 260 150 @("主流长片叙事结构", "角色、世界观与情节推进清晰", "画面服务于故事和情感代入") 16 | Out-Null
Add-CaseLine $s "观察点：它具有中国文化元素，但是否等同于实验动画？"

$s = $pres.Slides.Item(11)
Set-Title $s "案例对比：《小蝌蚪找妈妈》"
Clear-TextExcept $s @(3)
Add-Bullets $s 42 80 280 150 @("水墨材料语言进入动画", "运动、笔墨与中国画意境结合", "形式探索参与作品表达") 16 | Out-Null
Add-CaseLine $s "观察点：材料和形式如何改变动画的观看体验？"

# Slide 12: definition
$s = $pres.Slides.Item(12)
Set-TextById $s 3 "实验动画概述"
Set-TextById $s 5 "什么是实验动画？"
Set-TextById $s 6 "实验动画是一种强调探索性、创新性和个体表达的动画创作方式。它不以完整故事叙事作为唯一目标，而是重视主题、材料、叙事结构、视听表现形式与观看体验的实验。`r`r它通过非常规方法拓展动画表达边界，使动画不只是讲故事，也成为感知、观念和材料运动的表达。" 14.2
Set-TextById $s 9 "实验动画 = 探索 + 表达 + 形式/材料创新" 15

# Slide 13: definition and features
$s = $pres.Slides.Item(13)
Set-Title $s "实验动画定义特点"
Clear-TextExcept $s @(3)
Add-TextBox $s 60 82 255 32 "实验动画定义" 20 | Out-Null
Add-TextBox $s 60 125 255 120 "广义上指具有探索性和创新性的动画作品；狭义上强调创作者摆脱固定工业规范和市场目标，以个体观念、形式实验与媒介探索建立独特表达。" 15 | Out-Null
Add-TextBox $s 392 82 255 32 "实验动画特点" 20 | Out-Null
Add-Bullets $s 392 125 255 140 @("非单一叙事导向", "强调材料探索", "注重形式创新", "鼓励个体表达") 15 | Out-Null

# Development / case slides
$s = $pres.Slides.Item(14)
Set-Title $s "实验动画案例：形式与节奏"
Add-CaseLine $s "影片案例：《对角交响曲》（Viking Eggeling，1924）｜观察抽象形体、节奏与音乐性。"
$s = $pres.Slides.Item(15)
Set-Title $s "实验动画案例：真人逐格与社会寓言"
Add-CaseLine $s "影片案例：《邻居》（Norman McLaren，1952）｜观察真人逐格、反战寓言与运动实验。"
$s = $pres.Slides.Item(16)
Set-Title $s "中国动画实验传统案例"
Add-CaseLine $s "影片案例：《小蝌蚪找妈妈》（1960）｜观察水墨材料、笔触运动与民族化探索。"

# Slide 17: classification framework
$s = $pres.Slides.Item(17)
Set-TextById $s 3 "实验动画分类"
Set-TextById $s 4 "实验动画的分类框架"
Set-TextById $s 6 "三种观察维度"
Set-TextById $s 8 "主题"
Set-TextById $s 10 "创作方式"
Set-TextById $s 12 "材料"
Set-TextById $s 14 "情绪、心理、社会议题、文化记忆、哲学观念" 14.5
Set-TextById $s 16 "手绘、剪纸、偶动画、实拍逐格、数字合成" 14.5
Set-TextById $s 18 "水墨、纸张、布料、沙、实物、数字笔刷" 14.5

# Slides 18-22: classification examples
$s = $pres.Slides.Item(18)
Set-Title $s "按创作方式分类：传统手绘与二维实验"
Add-CaseLine $s "比较案例：《仿制品》《大闹天宫》｜观察手绘风格、角色运动与画面组织差异。"

$s = $pres.Slides.Item(19)
Set-Title $s "按创作方式分类：剪纸与皮影风格"
Add-CaseLine $s "比较案例：《猪八戒吃西瓜》《功夫熊猫2》皮影片段｜观察平面造型、镂空感与民间美术转译。"

$s = $pres.Slides.Item(20)
Set-Title $s "按创作方式分类：偶动画"
Add-CaseLine $s "比较案例：《阿凡提》《鬼妈妈》｜观察偶的材质、空间调度与逐格表演。"

$s = $pres.Slides.Item(21)
Set-Title $s "按创作方式分类：数字实验与混合技法"
Add-CaseLine $s "比较案例：《Please Say Something》《功夫兔》｜观察3D实验、实拍结合二维与媒介混合。"

$s = $pres.Slides.Item(22)
Set-Title $s "按材料分类：水墨材料语言"
Add-CaseLine $s "比较案例：《山水情》《小蝌蚪找妈妈》｜观察墨色浓淡、留白、笔触与运动节奏。"

# Slides 23-26: repurpose culture overview pages
$s = $pres.Slides.Item(23)
Clear-TextExcept $s @()
Add-TextBox $s 76 120 90 80 "03" 42 | Out-Null
Add-TextBox $s 180 124 450 78 "分类框架与案例辨析" 32 | Out-Null
Add-TextBox $s 182 205 460 42 "主题、创作方式、材料不是单选题，而是交叉的分析维度。" 16 | Out-Null

$s = $pres.Slides.Item(24)
Set-Title $s "按主题分类：比较案例"
Clear-TextExcept $s @(3)
Add-Bullets $s 56 82 610 165 @(
    "情绪与心理：《父与女》《回忆积木小屋》",
    "社会观察：《邻居》《种树的牧羊人》",
    "文化寓言：《九色鹿》《山水情》",
    "梦境与潜意识：抽象动画、非线性诗性短片"
) 17 | Out-Null
Add-TextBox $s 58 285 600 52 "观察问题：作品主要表达什么？它依靠故事事件、情绪经验，还是文化意象来组织观看？" 15 | Out-Null

$s = $pres.Slides.Item(25)
Set-Title $s "分类不是唯一归属"
Clear-TextExcept $s @(3)
Add-Bullets $s 58 84 610 180 @(
    "同一作品可同时从主题、创作方式和材料三个维度分析。",
    "例：《小蝌蚪找妈妈》既是水墨材料实验，也包含中国画审美和民族化动画探索。",
    "例：《猪八戒吃西瓜》可从剪纸材料、民间故事主题和装饰性造型三个角度理解。"
) 17 | Out-Null
Add-TextBox $s 58 298 600 44 "分析三问：题材讲什么？创作方式怎么组织？材料如何参与表达？" 16 | Out-Null

$s = $pres.Slides.Item(26)
Set-Title $s "中国动画学派的实验传统简述"
Clear-TextExcept $s @(3)
Add-Bullets $s 58 84 610 185 @(
    "水墨动画：以笔墨、留白和意境建立动画语言。",
    "剪纸动画：将民间美术造型、皮影感和平面运动结合。",
    "偶动画与实物动画：通过材料质感、空间调度和逐格表演形成独特表达。",
    "本课不展开系统中国文化史，重点理解中国动画传统中的形式实验和材料探索。"
) 16 | Out-Null

# Slides 38-41 updated before deleting old 27-37
$s = $pres.Slides.Item(38)
Clear-TextExcept $s @()
Add-TextBox $s 76 120 90 80 "04" 42 | Out-Null
Add-TextBox $s 180 124 450 78 "课堂检测与任务布置" 32 | Out-Null

$s = $pres.Slides.Item(39)
Set-Title $s "课堂测试与反馈"
Clear-TextExcept $s @(3)
Add-Bullets $s 54 82 620 220 @(
    "实验动画区别于主流动画的关键特征不包括哪一项？",
    "下列哪一项属于按材料分类的观察角度？",
    "一部作品是否只能归入一种实验动画类型？请说明理由。",
    "材料语言在实验动画中可以如何参与表达？"
) 17 | Out-Null

$s = $pres.Slides.Item(40)
Clear-TextExcept $s @()
Add-TextBox $s 76 120 90 80 "05" 42 | Out-Null
Add-TextBox $s 180 124 450 78 "课后实践" 32 | Out-Null
Add-TextBox $s 182 205 460 42 "材料语言测试：从材料出发，理解运动与表达。" 16 | Out-Null

$s = $pres.Slides.Item(41)
Set-Title $s "材料语言测试任务说明"
Clear-TextExcept $s @(3)
Add-Bullets $s 52 78 620 205 @(
    "利用纸张、布料、树叶等综合材料，或使用数字笔刷，制作一段约5秒的《材质运动测试》。",
    "不要求完整叙事，但需要体现清晰的实验意识。",
    "鼓励尝试非常规材料、纹理、节奏或视觉组织方式。",
    "提交测试视频及简短文字说明：材料、表现意图、遇到的问题。"
) 16 | Out-Null
Add-TextBox $s 54 310 610 42 "案例参考：水墨运动测试、剪纸形变测试、树叶/布料逐格测试、数字笔刷轨迹测试。" 15 | Out-Null

# Remove outdated broad China-culture/story-writing pages from this first lesson.
for ($i = 37; $i -ge 27; $i--) {
    $pres.Slides.Item($i).Delete()
}

$pres.Save()
$pres.Close()
$app.Quit()

Write-Output $out
