
            for field_code, field_value in liangti_data.items():
                # 获取字段的中文名称
                field_aliases = {
                    'hips': ['hip', '臀围', '全臀围', '净臀围', 'Hip', 'Full Hip'],
                    'hips': ['fullHip', 'fullHipWidth', '全臀围', 'Full Hip'],
                    'hip': ['hips', '臀围', '全臀围', '净臀围', 'Hip', 'Full Hip'],
                }
                mapped_info = liangti_mapping.get(field_code, {})
                    field_candidates.append(english_name.strip())
                for alias in field_aliases.get(field_code, []):
                    field_candidates.append(alias)
                if field_name.startswith('全') and len(field_name) > 1:
                if field_code != 'hips' and field_name.startswith('全') and len(field_name) > 1:
                    field_candidates.append(field_name[1:])
                if field_name.startswith('净') and len(field_name) > 1:
                    field_candidates.append(field_name[1:])
                        
                        const getText = (el) => (el ? (el.innerText || el.textContent || '').trim() : '');
                        const normalizeText = (text) => (text || '').replace(/\\s+/g, '').replace(/[()（）]/g, '').trim();
                        const normalizedCandidates = fieldCandidates.map(normalizeText).filter(Boolean);
                        const normalizedCandidates = Array.from(new Set(fieldCandidates.map(normalizeText).filter(Boolean)));
                        const textMatches = (text) => {
                            const normalizedText = normalizeText(text);
                            if (!normalizedText) return false;
                            return normalizedCandidates.some(candidate =>
                                normalizedText === candidate ||
                                normalizedText.includes(candidate) ||
                                candidate.includes(normalizedText)
                            );
                            return normalizedCandidates.some(candidate => normalizedText === candidate);
                        };
                        const rowPartCellMatches = (row) => {
                            const cells = Array.from(row.querySelectorAll('td, th'));
                            return cells.find(cell => textMatches(getText(cell))) || null;
                        };
                        const isVisible = (el) => {
                            if (!el) return false;
                            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
                        };
                        const findColumnIndex = (headers, keywords) => {
                            const normalizedKeywords = keywords.map(normalizeText);
                            for (let i = 0; i < headers.length; i++) {
                                const text = getText(headers[i]).replace(/\\s+/g, '');
                                const text = normalizeText(getText(headers[i]));
                                console.log('表头:', i, text);
                                if (keywords.some(keyword => text.includes(keyword))) return i;
                                if (text.includes('特体部位')) continue;
                                if (normalizedKeywords.some(keyword => text === keyword)) return i;
                            }
                            for (let i = 0; i < headers.length; i++) {
                                const text = normalizeText(getText(headers[i]));
                                if (text.includes('特体部位')) continue;
                                if (normalizedKeywords.some(keyword => text.includes(keyword))) return i;
                            }
                            return -1;
                        };
                                if (!rowText) continue;
                                const normalizedRowText = normalizeText(rowText);
                                if (normalizedRowText) visiblePartNames.push(normalizedRowText);
                                if (textMatches(rowText)) {
                                const matchedPartCell = rowPartCellMatches(row);
                                if (matchedPartCell) {
                                    targetRow = row;
                                    const cells = row.querySelectorAll('td, th');
                                    targetCell = Array.from(cells).find(cell => cell.querySelector('.el-input-number__increase, .el-input-number__decrease, input')) || null;
                                    targetCell = Array.from(cells).find(cell => cell !== matchedPartCell && cell.querySelector('.el-input-number__increase, .el-input-number__decrease, input')) || null;
                                    console.log('✓ 通过整行文本找到目标行:', normalizedRowText);
                                    break;
                                }
