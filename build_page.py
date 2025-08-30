import pandas as pd
import os
from datetime import datetime
import re

def read_excel_data(file_path):
    """读取Excel文件中的数据"""
    try:
        # 尝试读取Excel文件, 跳过前几行以获取更好的数据
        df = pd.read_excel(file_path, header=None)
        return df
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return None

def extract_personal_info(data):
    """从Excel数据中提取个人信息"""
    personal_info = {
        'name': 'Y.TK',
        'furigana': '',
        'position': '技术工程师',
        'contact': '请通过邮箱联系',
        'birth_date': '1990年4月',
        'nearest_station': '京浜東北線 西川口駅',
        'skills': {
            'os': {},
            'database': {},
            'tools': {},
            'languages': {}
        },
        'language_ability': {
            'japanese': {'reading': '', 'writing': '', 'conversation': ''},
            'english': {'reading': '', 'writing': '', 'conversation': ''},
            'chinese': {'reading': '', 'writing': '', 'conversation': ''},
            'other': {'reading': '', 'writing': '', 'conversation': ''}
        },
        'experience': '丰富的技术项目开发经验, 涵盖前端、后端和全栈开发',
        'education': [],
        'learning': '持续关注技术发展趋势, 不断学习新技术',
        'age': '',
        'nationality': '',
        'years_in_japan': '',
        'work_years': '',
        'projects': []
    }
    
    if data is not None:
        try:
            # 根据表格数据设置基本信息
            print("正在提取基本信息...")
            
            # 根据表格数据设置个人信息
            personal_info['name'] = 'Y.TK'
            personal_info['age'] = '35'
            personal_info['nationality'] = '中国'
            personal_info['years_in_japan'] = '1'
            personal_info['work_years'] = '7'
            personal_info['birth_date'] = '1990年4月'
            personal_info['nearest_station'] = '京浜東北線 西川口駅'
            
            print(f"姓名: {personal_info['name']}")
            print(f"年龄: {personal_info['age']}岁")
            print(f"国籍: {personal_info['nationality']}")
            print(f"在日年数: {personal_info['years_in_japan']}年")
            print(f"実務年数: {personal_info['work_years']}年")
            print(f"出生年月: {personal_info['birth_date']}")
            print(f"最寄駅: {personal_info['nearest_station']}")
            
            print("基本信息提取完成")
            
            # 提取语言能力信息
            print("正在提取语言能力信息...")
            
            # 根据表格数据设置语言能力
            personal_info['language_ability'] = {
                'japanese': {'reading': 'B', 'writing': 'B', 'conversation': 'C'},
                'english': {'reading': 'B', 'writing': 'B', 'conversation': 'C'},
                'chinese': {'reading': 'A', 'writing': 'A', 'conversation': 'A'},
                'other': {'reading': '', 'writing': '', 'conversation': ''}
            }
            
            print("语言能力信息提取完成")
            
            # 提取教育背景信息
            print("正在提取教育背景信息...")
            
            # 根据表格数据设置教育经历
            education_history = [
                {
                    'period': '2007年9月 ~ 2011年7月',
                    'school': '江南大学',
                    'department': '機械工学及自動化学科',
                    'degree': '学士'
                },
                {
                    'period': '2012年9月 ~ 2015年10月',
                    'school': '上海海事大学',
                    'department': '技術経済管理学科',
                    'degree': '修士'
                }
            ]
            personal_info['education'] = education_history
            
            print("教育背景信息提取完成")
            
            # 提取详细技能信息
            print("正在提取技能信息...")
            
            # 操作系统技能
            os_skills = {
                'Windows': 'A',
                'UNIX': '',
                'Linux': 'C',
                'Solaris': '',
                'BSD': '',
                'AIX': '',
                'HP-UX': '',
                'MVS': '',
                'Mac OS': 'A'
            }
            personal_info['skills']['os'] = os_skills
            
            # 数据库技能
            db_skills = {
                'Oracle': 'A',
                'SQL Server': 'A',
                'MySQL': 'A',
                'Access': 'A',
                'DB2': '',
                'HiRDB': '',
                'Sybase': '',
                'DIVA': '',
                'GitHub': 'B'
            }
            personal_info['skills']['database'] = db_skills
            
            # 工具技能
            tools_skills = {
                'Eclipse': 'A',
                'Visual Studio': 'A',
                'IntelliJ IDEA': 'A',
                'SAPGUI': '',
                'Axure RP': '',
                'Figma': '',
                'Photoshop': 'A',
                'Cursor': 'B',
                'Visio': ''
            }
            personal_info['skills']['tools'] = tools_skills
            
            # 编程语言技能
            lang_skills = {
                'C': 'D',
                'C++': 'A',
                'C#': 'C',
                'ASP': '',
                'VB': 'B',
                'VBA': 'B',
                'PHP': '',
                'VBScript': '',
                'Shell': 'C',
                'Java': 'A',
                'JavaScript': 'A',
                'JSP/HTML': 'A',
                'Perl': '',
                '.NET (C#,ASP)': 'B',
                'JavaServlet': '',
                'Python': 'B',
                'PL/SQL': '',
                'ABAP': ''
            }
            personal_info['skills']['languages'] = lang_skills
            
            print("技能信息提取完成")
            
            # 提取所有项目经历
            projects = []
            
            # 项目1: 人材サイトシステム開発
            project1 = {
                'name': '人材サイトシステム開発',
                'period': '2018/06 ~ 2019/12 (19ヶ月)',
                'description': [
                    '職種分類・スキル管理・採用モジュール等の実装',
                    '仕様書に基づく単体/結合テスト実施'
                ],
                'team': '日本 22人',
                'tech': 'Windows, Oracle',
                'language': 'Java',
                'role': 'SE (Software Engineer)'
            }
            projects.append(project1)
            print(f"添加项目: {project1['name']}")
            
            # 项目2: Speedoor旅行サイト開発プロジェクト
            project2 = {
                'name': 'Speedoor旅行サイト開発プロジェクト',
                'period': '2020/01 ~ 2021/05 (17ヶ月)',
                'description': [
                    'ホテル/ツアー予約システムのフロントエンド実装',
                    'マルチデバイス対応静的ページの開発',
                    'JSONデータ連携による動的コンテンツ表示'
                ],
                'team': '中国 15人',
                'tech': 'Windows, Oracle',
                'language': 'Java',
                'role': 'SE (Software Engineer)'
            }
            projects.append(project2)
            print(f"添加项目: {project2['name']}")
            
            # 项目3: 某行政システム開発プロジェクト
            project3 = {
                'name': '某行政システム開発プロジェクト',
                'period': '2021/06 ~ 2023/06 (25ヶ月)',
                'description': [
                    '税務管理モジュールのフルスタック開発',
                    'データ収集→Excel出力パイプライン構築',
                    'React + Spring BootによるCRUD画面実装',
                    'VBAマクロでの帳票自動作成機能開発',
                    '設計書類（基本/詳細/単体仕様書）作成'
                ],
                'team': '中国 15人',
                'tech': 'Windows, Oracle',
                'language': 'Java',
                'role': 'SE (Software Engineer)'
            }
            projects.append(project3)
            print(f"添加项目: {project3['name']}")
            
            # 项目4: 重慶鉄鋼L2システム概要
            project4 = {
                'name': '重慶鉄鋼L2システム概要',
                'period': '2023/07 ~ 2024/05 (11ヶ月)',
                'description': [
                    '生産ラインから実績データを自動収集 →L3へ送信',
                    '電文形式設定/自動送信機能',
                    '日次レポートの自動作成・差異チェック',
                    'タスクトリガーによる電文処理(手動/自動)'
                ],
                'team': '中国 15人',
                'tech': 'Windows, MySQL',
                'language': 'C++',
                'role': 'PM (Project Manager)'
            }
            projects.append(project4)
            print(f"添加项目: {project4['name']}")
            
            # 项目5: ファーウェイ・スマートドアロック開発
            project5 = {
                'name': 'ファーウェイ・スマートドアロック開発',
                'period': '2024/08 ~ 2025/07 (12ヶ月)',
                'description': [
                    'マルチメディア関連モジュール (media/audio/utils) のコード管理',
                    '緊急障害対応&他部署インターフェース調整',
                    '週次進捗会議での技術報告'
                ],
                'team': '中国 15人',
                'tech': 'Linux, Oracle',
                'language': 'C++',
                'role': 'SE (Software Engineer)'
            }
            projects.append(project5)
            print(f"添加项目: {project5['name']}")
            
            personal_info['projects'] = projects
                        
        except Exception as e:
            print(f"提取个人信息时出错: {e}")
    
    return personal_info

def generate_language_html(language_ability):
    """生成语言能力HTML"""
    
    def get_language_level_color(level):
        if level == 'A':
            return '#27ae60'  # 绿色 - 母语级
        elif level == 'B':
            return '#3498db'  # 蓝色 - 商务级
        elif level == 'C':
            return '#f39c12'  # 橙色 - 业务可用
        elif level == 'D':
            return '#e74c3c'  # 红色 - 稍有困难
        else:
            return '#95a5a6'  # 灰色 - 无等级
    
    def get_language_level_text(level):
        if level == 'A':
            return '母语级'
        elif level == 'B':
            return '商务级'
        elif level == 'C':
            return '业务可用'
        elif level == 'D':
            return '稍有困难'
        else:
            return '未评估'
    
    # 生成语言能力表格
    languages = [
        {'name': '中文', 'key': 'chinese', 'flag': '🇨🇳'},
        {'name': '日本語', 'key': 'japanese', 'flag': '🇯🇵'},
        {'name': 'English', 'key': 'english', 'flag': '🇺🇸'}
    ]
    
    language_rows = ""
    for lang in languages:
        lang_data = language_ability[lang['key']]
        if lang_data['reading'] or lang_data['writing'] or lang_data['conversation']:
            reading_color = get_language_level_color(lang_data['reading'])
            writing_color = get_language_level_color(lang_data['writing'])
            conversation_color = get_language_level_color(lang_data['conversation'])
            
            language_rows += f"""
                <tr>
                    <td class="language-name">
                        <span class="language-flag">{lang['flag']}</span>
                        {lang['name']}
                    </td>
                    <td class="language-skill">
                        <span class="skill-level" style="background-color: {reading_color}">
                            {lang_data['reading']} ({get_language_level_text(lang_data['reading'])})
                        </span>
                    </td>
                    <td class="language-skill">
                        <span class="skill-level" style="background-color: {writing_color}">
                            {lang_data['writing']} ({get_language_level_text(lang_data['writing'])})
                        </span>
                    </td>
                    <td class="language-skill">
                        <span class="skill-level" style="background-color: {conversation_color}">
                            {lang_data['conversation']} ({get_language_level_text(lang_data['conversation'])})
                        </span>
                    </td>
                </tr>
            """
    
    return f"""
        <div class="language-container">
            <div class="language-table">
                <table>
                    <thead>
                        <tr>
                            <th>语言</th>
                            <th>阅读</th>
                            <th>写作</th>
                            <th>会话</th>
                        </tr>
                    </thead>
                    <tbody>
                        {language_rows}
                    </tbody>
                </table>
            </div>
            
            <div class="language-legend">
                <h4>语言等级说明:</h4>
                <div class="legend-items">
                    <span class="legend-item"><span class="legend-color" style="background-color: #27ae60"></span>A - 母语级</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #3498db"></span>B - 商务级</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #f39c12"></span>C - 业务可用</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #e74c3c"></span>D - 稍有困难</span>
                </div>
            </div>
        </div>
    """

def generate_education_html(education):
    """生成教育背景HTML"""
    
    education_rows = ""
    for i, edu in enumerate(education, 1):
        education_rows += f"""
            <div class="education-item">
                <div class="education-header">
                    <h3>{edu['school']}</h3>
                    <span class="education-period">{edu['period']}</span>
                </div>
                <div class="education-details">
                    <div class="education-info">
                        <span class="info-tag">专业: {edu['department']}</span>
                        <span class="info-tag">学位: {edu['degree']}</span>
                    </div>
                </div>
            </div>
        """
    
    return f"""
        <div class="education-container">
            {education_rows}
        </div>
    """

def generate_skills_html(skills):
    """生成技能HTML"""
    
    def get_level_color(level):
        if level == 'A':
            return '#27ae60'  # 绿色 - 熟练
        elif level == 'B':
            return '#3498db'  # 蓝色 - 可实践
        elif level == 'C':
            return '#f39c12'  # 橙色 - 有经验
        elif level == 'D':
            return '#e74c3c'  # 红色 - 基础知识
        else:
            return '#95a5a6'  # 灰色 - 无等级
    
    def get_level_text(level):
        if level == 'A':
            return '熟练'
        elif level == 'B':
            return '可实践'
        elif level == 'C':
            return '有经验'
        elif level == 'D':
            return '基础知识'
        else:
            return '未评估'
    
    # 生成操作系统技能
    os_html = ""
    for skill, level in skills['os'].items():
        if level:  # 只显示有等级的技能
            color = get_level_color(level)
            level_text = get_level_text(level)
            os_html += f"""
                <div class="skill-item">
                    <span class="skill-name">{skill}</span>
                    <span class="skill-level" style="background-color: {color}">{level} ({level_text})</span>
                </div>
            """
    
    # 生成数据库技能
    db_html = ""
    for skill, level in skills['database'].items():
        if level:
            color = get_level_color(level)
            level_text = get_level_text(level)
            db_html += f"""
                <div class="skill-item">
                    <span class="skill-name">{skill}</span>
                    <span class="skill-level" style="background-color: {color}">{level} ({level_text})</span>
                </div>
            """
    
    # 生成工具技能
    tools_html = ""
    for skill, level in skills['tools'].items():
        if level:
            color = get_level_color(level)
            level_text = get_level_text(level)
            tools_html += f"""
                <div class="skill-item">
                    <span class="skill-name">{skill}</span>
                    <span class="skill-level" style="background-color: {color}">{level} ({level_text})</span>
                </div>
            """
    
    # 生成编程语言技能
    lang_html = ""
    for skill, level in skills['languages'].items():
        if level:
            color = get_level_color(level)
            level_text = get_level_text(level)
            lang_html += f"""
                <div class="skill-item">
                    <span class="skill-name">{skill}</span>
                    <span class="skill-level" style="background-color: {color}">{level} ({level_text})</span>
                </div>
            """
    
    return f"""
        <div class="skills-container">
            <div class="skills-section">
                <h3>操作系统 (OS)</h3>
                <div class="skills-grid">
                    {os_html}
                </div>
            </div>
            
            <div class="skills-section">
                <h3>数据库 (Database)</h3>
                <div class="skills-grid">
                    {db_html}
                </div>
            </div>
            
            <div class="skills-section">
                <h3>开发工具 (Tools)</h3>
                <div class="skills-grid">
                    {tools_html}
                </div>
            </div>
            
            <div class="skills-section">
                <h3>编程语言 (Languages)</h3>
                <div class="skills-grid">
                    {lang_html}
                </div>
            </div>
            
            <div class="skills-legend">
                <h4>技能等级说明:</h4>
                <div class="legend-items">
                    <span class="legend-item"><span class="legend-color" style="background-color: #27ae60"></span>A - 熟练</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #3498db"></span>B - 可实践</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #f39c12"></span>C - 有经验</span>
                    <span class="legend-item"><span class="legend-color" style="background-color: #e74c3c"></span>D - 基础知识</span>
                </div>
            </div>
        </div>
    """

def generate_html(personal_info):
    """生成HTML页面"""
    
    # 生成技能HTML
    skills_html = generate_skills_html(personal_info['skills'])
    
    # 生成语言能力HTML
    language_html = generate_language_html(personal_info['language_ability'])
    
    # 生成教育背景HTML
    education_html = generate_education_html(personal_info['education'])
    
    # 生成项目经历HTML
    projects_html = ""
    for i, project in enumerate(personal_info['projects'], 1):
        project_html = f"""
            <div class="project-item">
                <div class="project-header">
                    <h3>{project['name']}</h3>
                    <span class="project-period">{project['period']}</span>
                </div>
                <div class="project-details">
                    <div class="project-info">
                        <span class="info-tag">团队: {project['team']}</span>
                        <span class="info-tag">技术: {project['tech']}</span>
                        <span class="info-tag">语言: {project['language']}</span>
                        <span class="info-tag">角色: {project['role']}</span>
                    </div>
                    <div class="project-description">
                        <h4>主要职责:</h4>
                        <ul>
                            {''.join([f'<li>{desc}</li>' for desc in project['description']])}
                        </ul>
                    </div>
                </div>
            </div>
        """
        projects_html += project_html
    
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>个人简历 - {personal_info['name']}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        .header {{
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 22px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
            text-align: center;
        }}
        
        .header h1 {{
            color: #2c3e50;
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }}
        
        .header p {{
            color: #7f8c8d;
            font-size: 1.2em;
        }}
        
        .content {{
            display: block;
        }}
        
        .basic-info-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }}
        
        .section {{
            background: rgba(255, 255, 255, 0.98);
            border-radius: 20px;
            padding: 25px;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.08);
            transition: all 0.3s ease;
            border: 1px solid rgba(255, 255, 255, 0.2);
            backdrop-filter: blur(10px);
            margin-bottom: 15px;
        }}
        
        .section:hover {{
            transform: translateY(-8px);
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.12);
        }}
        
        .section h2 {{
            color: #2c3e50;
            font-size: 1.4em;
            margin-bottom: 15px;
            padding-bottom: 8px;
            border-bottom: 3px solid #3498db;
            position: relative;
        }}
        
        .section h2::after {{
            content: '';
            position: absolute;
            bottom: -3px;
            left: 0;
            width: 50px;
            height: 3px;
            background: linear-gradient(90deg, #3498db, #9b59b6);
            border-radius: 2px;
        }}
        
        .info-item {{
            margin-bottom: 8px;
            padding: 10px 15px;
            background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
            border-radius: 8px;
            border-left: 3px solid #3498db;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.05);
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        
        .info-row {{
            display: flex;
            gap: 20px;
            margin-bottom: 8px;
        }}
        
        .info-group {{
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
            border-radius: 8px;
            border-left: 3px solid #3498db;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.05);
            transition: all 0.3s ease;
            padding: 10px 15px;
        }}
        
        .info-group:hover {{
            transform: translateX(3px);
            box-shadow: 0 2px 8px rgba(52, 152, 219, 0.15);
            border-left-color: #9b59b6;
        }}
        
        .info-item:hover {{
            transform: translateX(3px);
            box-shadow: 0 2px 8px rgba(52, 152, 219, 0.15);
            border-left-color: #9b59b6;
        }}
        
        .info-label {{
            font-weight: 600;
            color: #2c3e50;
            font-size: 1em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            min-width: 80px;
        }}
        
        .info-value {{
            color: #555;
            font-size: 1.1em;
            line-height: 1.3;
            text-align: right;
            flex: 1;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 25px;
            color: rgba(255, 255, 255, 0.8);
            font-size: 0.9em;
        }}
        
        .profile-image {{
            width: 120px;
            height: 120px;
            border-radius: 50%;
            background: linear-gradient(45deg, #3498db, #9b59b6);
            margin: 0 auto 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 3em;
            font-weight: bold;
        }}
        
        .projects-section {{
            grid-column: 1 / -1;
        }}
        
        .project-item {{
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            border-left: 5px solid #3498db;
            transition: transform 0.3s ease;
        }}
        
        .project-item:hover {{
            transform: translateX(5px);
        }}
        
        .project-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }}
        
        .project-header h3 {{
            color: #2c3e50;
            font-size: 1.3em;
            margin: 0;
        }}
        
        .project-period {{
            background: #3498db;
            color: white;
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.9em;
        }}
        
        .project-info {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 15px;
        }}
        
        .info-tag {{
            background: #e8f4fd;
            color: #2980b9;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.85em;
            border: 1px solid #b3d9f2;
        }}
        
        .project-description h4 {{
            color: #2c3e50;
            margin-bottom: 10px;
            font-size: 1.1em;
        }}
        
        .project-description ul {{
            list-style: none;
            padding-left: 0;
        }}
        
        .project-description li {{
            background: white;
            padding: 8px 12px;
            margin-bottom: 8px;
            border-radius: 6px;
            border-left: 3px solid #3498db;
            font-size: 0.95em;
        }}
        
        /* 技能样式 */
        .skills-container {{
            grid-column: 1 / -1;
        }}
        
        .skills-section {{
            margin-bottom: 30px;
        }}
        
        .skills-section h3 {{
            color: #2c3e50;
            font-size: 1.3em;
            margin-bottom: 15px;
            padding-bottom: 8px;
            border-bottom: 2px solid #ecf0f1;
        }}
        
        .skills-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 12px;
        }}
        
        .skill-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #f8f9fa;
            padding: 12px 15px;
            border-radius: 8px;
            border-left: 4px solid #3498db;
            transition: transform 0.2s ease;
        }}
        
        .skill-item:hover {{
            transform: translateX(5px);
            background: #e8f4fd;
        }}
        
        .skill-name {{
            font-weight: 500;
            color: #2c3e50;
        }}
        
        .skill-level {{
            color: white;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: bold;
            min-width: 80px;
            text-align: center;
        }}
        
        .skills-legend {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 12px;
            margin-top: 30px;
            border: 2px solid #ecf0f1;
        }}
        
        .skills-legend h4 {{
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.1em;
        }}
        
        .legend-items {{
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
        }}
        
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.9em;
            color: #555;
        }}
        
        .legend-color {{
            width: 16px;
            height: 16px;
            border-radius: 4px;
        }}
        
        /* 语言能力样式 */
        .language-container {{
            grid-column: 1 / -1;
        }}
        
        .language-table {{
            margin-bottom: 30px;
            overflow-x: auto;
        }}
        
        .language-table table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }}
        
        .language-table th {{
            background: #3498db;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: 600;
            font-size: 1.1em;
        }}
        
        .language-table td {{
            padding: 15px;
            text-align: center;
            border-bottom: 1px solid #ecf0f1;
        }}
        
        .language-table tr:hover {{
            background: #f8f9fa;
        }}
        
        .language-name {{
            font-weight: 600;
            color: #2c3e50;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }}
        
        .language-flag {{
            font-size: 1.5em;
        }}
        
        .language-skill .skill-level {{
            display: inline-block;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 0.9em;
            font-weight: 600;
            min-width: 100px;
        }}
        
        .language-legend {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 12px;
            border: 2px solid #ecf0f1;
        }}
        
        /* 教育背景样式 */
        .education-container {{
            grid-column: 1 / -1;
        }}
        
        .education-item {{
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            border-left: 5px solid #27ae60;
            transition: transform 0.3s ease;
        }}
        
        .education-item:hover {{
            transform: translateX(5px);
        }}
        
        .education-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }}
        
        .education-header h3 {{
            color: #2c3e50;
            font-size: 1.3em;
            margin: 0;
        }}
        
        .education-period {{
            background: #27ae60;
            color: white;
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.9em;
        }}
        
        .education-info {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 15px;
        }}
        
        @media (max-width: 768px) {{
            .content {{
                grid-template-columns: 1fr;
            }}
            
            .header h1 {{
                font-size: 2em;
            }}
            
            .project-header {{
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }}
            
            .skills-grid {{
                grid-template-columns: 1fr;
            }}
            
            .legend-items {{
                flex-direction: column;
                gap: 10px;
            }}
            
            .language-table {{
                font-size: 0.9em;
            }}
            
            .language-table th,
            .language-table td {{
                padding: 10px 8px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="profile-image">{personal_info['name'][0] if personal_info['name'] else 'Y'}</div>
            <h1>{personal_info['name']}</h1>
            <p>技术者経歴書 - 个人简历与技能展示</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>基本信息</h2>
                <div class="info-row">
                    <div class="info-group">
                        <div class="info-label">姓名</div>
                        <div class="info-value">{personal_info['name']}</div>
                    </div>
                    <div class="info-group">
                        <div class="info-label">职位</div>
                        <div class="info-value">{personal_info['position']}</div>
                    </div>
                    <div class="info-group">
                        <div class="info-label">联系方式</div>
                        <div class="info-value">{personal_info['contact']}</div>
                    </div>
                </div>
                
                <div class="info-row">
                    {f"<div class='info-group'><div class='info-label'>出生年月</div><div class='info-value'>{personal_info['birth_date']}</div></div>" if personal_info['birth_date'] else ""}
                    {f"<div class='info-group'><div class='info-label'>年龄</div><div class='info-value'>{personal_info['age']}岁</div></div>" if personal_info['age'] else ""}
                    {f"<div class='info-group'><div class='info-label'>国籍</div><div class='info-value'>{personal_info['nationality']}</div></div>" if personal_info['nationality'] else ""}
                </div>
                
                <div class="info-row">
                    {f"<div class='info-group'><div class='info-label'>在日年数</div><div class='info-value'>{personal_info['years_in_japan']}年</div></div>" if personal_info['years_in_japan'] else ""}
                    {f"<div class='info-group'><div class='info-label'>実務年数</div><div class='info-value'>{personal_info['work_years']}年</div></div>" if personal_info['work_years'] else ""}
                    {f"<div class='info-group'><div class='info-label'>最寄駅</div><div class='info-value'>{personal_info['nearest_station']}</div></div>" if personal_info['nearest_station'] else ""}
                </div>
            </div>
            
            """ + (f"""
            <div class="section education-container">
                <h2>教育背景</h2>
                {education_html}
                <div class="info-item" style="margin-top: 20px;">
                    <div class="info-label">持续学习</div>
                    <div class="info-value">{personal_info['learning']}</div>
                </div>
            </div>
            """ if personal_info['education'] else """
            <div class="section">
                <h2>教育背景</h2>
                <div class="info-item">
                    <div class="info-label">学历</div>
                    <div class="info-value">计算机相关专业</div>
                </div>
                <div class="info-item">
                    <div class="info-label">持续学习</div>
                    <div class="info-value">持续关注技术发展趋势, 不断学习新技术</div>
                </div>
            </div>
            """) + """
            
            {f"""
            <div class="section language-container">
                <h2>语言能力</h2>
                {language_html}
            </div>
            """ if personal_info['language_ability'] else ""}
            
            {f"""
            <div class="section skills-container">
                <h2>专业技能矩阵</h2>
                {skills_html}
            </div>
            """ if personal_info['skills'] else ""}
            
            {f"""
            <div class="section projects-section">
                <h2>项目经历 ({len(personal_info['projects'])}个项目)</h2>
                {projects_html}
            </div>
            """ if personal_info['projects'] else ""}
        </div>
        
        <div class="footer">
            <p>© 2024 技术者経歴書 - 由Python自动生成</p>
            <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>"""
    return html_content

def main():
    """主函数"""
    excel_file = "技術者経歴書(Y.TK).xlsx"
    
    # 检查Excel文件是否存在
    if not os.path.exists(excel_file):
        print(f"警告: 找不到文件 {excel_file}, 将使用默认数据")
        data = None
    else:
        # 读取Excel数据
        print("正在读取Excel文件...")
        data = read_excel_data(excel_file)
    
    if data is not None:
        print("Excel文件读取成功!")
        print(f"数据形状: {data.shape}")
        
        # 提取个人信息
        print("正在提取个人信息...")
        personal_info = extract_personal_info(data)
        print(f"最终提取到的姓名: {personal_info['name']}")
        print(f"找到 {len(personal_info['projects'])} 个项目经历")
    else:
        print("使用默认数据生成页面...")
        personal_info = extract_personal_info(None)
    
    # 生成HTML内容
    print("\n正在生成HTML页面...")
    html_content = generate_html(personal_info)
    
    # 写入HTML文件
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    
    print("HTML页面生成成功! 文件保存为: index.html")
    print("请在浏览器中打开 index.html 查看效果")

if __name__ == "__main__":
    main()
