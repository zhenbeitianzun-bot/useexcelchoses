        async function exportExcel() {
            try {
                if (currentResult.length === 0) {
                    alert('请先抽取人员');
                    return;
                }
                
                // 获取表单数据
                const projectName = document.getElementById('projectName').value;
                const applicantName = document.getElementById('applicantName').value;
                const supervisorName = document.getElementById('supervisorName').value;
                const applyDepartment = document.getElementById('applyDepartment').value;
                const peopleCount = document.getElementById('peopleCount').value;
                const selectedPeopleList = currentResult.join(', ');
                
                // 读取模板文件
                const response = await fetch('评审人员项目库随机抽取记录表.xls');
                if (!response.ok) {
                    throw new Error('模板文件加载失败');
                }
                const templateBuffer = await response.arrayBuffer();
                
                // 使用SheetJS读取.xls文件，启用所有解析选项
                const workbook = XLSX.read(templateBuffer, { 
                    type: 'array',
                    cellDates: true,
                    cellNF: true,
                    cellStyles: true,
                    cellHTML: true,
                    raw: false
                });
                
                // 获取第一个工作表
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // 打印工作表信息以便调试
                console.log('工作表:', worksheet);
                
                // 为所有单元格应用统一格式
                for (const cellRef in worksheet) {
                    if (cellRef.startsWith('!')) continue;
                    
                    const cell = worksheet[cellRef];
                    
                    // 设置统一格式
                    if (!cell.s) cell.s = {};
                    
                    // 设置字体：宋体，20号
                    cell.s.font = {
                        name: '宋体',
                        sz: 20
                    };
                    
                    // 设置对齐：水平居中，垂直居中，自动换行
                    cell.s.alignment = {
                        horizontal: 'center',
                        vertical: 'center',
                        wrapText: true
                    };
                    
                    // 添加边框
                    cell.s.border = {
                        top: { style: 'thin' },
                        right: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' }
                    };
                }
                
                // 直接指定单元格位置填充数据
                // 采购项目值在B3
                worksheet['B3'] = { v: projectName };
                worksheet['B3'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                // 申请人值在B4
                worksheet['B4'] = { v: applicantName };
                worksheet['B4'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                // 申请部门值在D4
                worksheet['D4'] = { v: applyDepartment };
                worksheet['D4'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                // 抽取人员数量值在B5
                worksheet['B5'] = { v: peopleCount };
                worksheet['B5'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                // 抽取人员名单值在D5
                worksheet['D5'] = { v: selectedPeopleList };
                worksheet['D5'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                // 监督人员值在B6
                worksheet['B6'] = { v: supervisorName };
                worksheet['B6'].s = {
                    font: { name: '宋体', sz: 20 },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border: { 
                        top: { style: 'thin' }, 
                        right: { style: 'thin' }, 
                        bottom: { style: 'thin' }, 
                        left: { style: 'thin' } 
                    }
                };
                
                console.log('所有数据已填充完成');
                console.log('采购项目:', projectName, '已填充到B3');
                console.log('申请人:', applicantName, '已填充到B4');
                console.log('申请部门:', applyDepartment, '已填充到D4');
                console.log('抽取人员数量:', peopleCount, '已填充到B5');
                console.log('抽取人员名单:', selectedPeopleList, '已填充到D5');
                console.log('监督人员:', supervisorName, '已填充到B6');
                
                // 导出修改后的文件
                const excelBuffer = XLSX.write(workbook, { 
                    bookType: 'xlsx', 
                    type: 'array',
                    cellDates: true,
                    cellNF: true,
                    cellStyles: true
                });
                
                // 创建Blob并下载
                const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = '评审人员项目库随机抽取记录表.xlsx';
                document.body.appendChild(a);
                a.click();
                
                // 清理
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                alert('Excel导出成功');
            } catch (error) {
                console.error('Excel导出失败:', error);
                alert('导出失败: ' + error.message);
            }
        }