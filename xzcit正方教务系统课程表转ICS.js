// ==UserScript==
// @name         xzcit正方教务系统课程表转ICS
// @namespace    http://tampermonkey.net/
// @version      1.2.4
// @description  正方教务系统课表导出 ICS（支持实验课，地点多地点显示）
// @author       CitrusCandy
// @match        *://jwpt.xzcit.cn/*
// @run-at       document-end
// @license      MIT
// ==/UserScript==

(function () {
    'use strict';

    const COURSE_URL = "kbcx/xskbcx_cxXskbcxIndex.html";
    const TZID = "Asia/Shanghai";
    const CRLF = "\r\n";

    const weekDayMap = {
        "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7
    };

    /* ================= 页面入口 ================= */
    window.addEventListener("load", () => {
        if (location.href.includes(COURSE_URL)) injectUI();
    });

    /* ================= UI 注入 ================= */
    function injectUI() {
        const toolbar = document.querySelector(".btn-toolbar.pull-right");
        if (!toolbar) {
            console.log("【UI】未找到工具栏");
            return;
        }

        // 检查是否已经添加过按钮
        if (toolbar.querySelector("#export-ics-btn")) {
            console.log("【UI】按钮已存在");
            return;
        }

        const btn = document.createElement("button");
        btn.id = "export-ics-btn";
        btn.className = "btn btn-default";
        btn.textContent = "生成 ICS";

        const dateInput = document.createElement("input");
        dateInput.id = "semester-start-date";
        dateInput.type = "date";
        dateInput.style.marginLeft = "10px";
        dateInput.value = getDefaultSemesterDate();

        toolbar.appendChild(btn);
        toolbar.appendChild(dateInput);

        btn.addEventListener("click", () => {
            const startDate = dateInput.value;
            const tableData = parseTable();
            const courses = parseCourses(tableData);

            // console.log("【DEBUG】课程总数:", courses.length);
            // console.log("【DEBUG】课程样例:", courses.slice(0, 3));

            if (courses.length === 0) {
                alert("未找到课程数据，请确认是否在课程表页面");
                return;
            }

            try {
                generateCalendar(courses, startDate);
                alert("ICS 文件已生成并下载");
            } catch (error) {
                console.error("生成日历时出错:", error);
                alert("生成ICS文件时出错，请查看控制台日志");
            }
        });
    }

    function getDefaultSemesterDate() {
        const now = new Date();
        const y = now.getFullYear();
        const m = now.getMonth() + 1;

        const targetMonth = (m >= 3 && m <= 8) ? 2 : 9;
        const d = new Date(y, targetMonth - 1, 1);
        const day = d.getDay() || 7;
        d.setDate(d.getDate() + (day === 1 ? 0 : 8 - day));

        return d.toISOString().slice(0, 10);
    }

    /* ================= 表格解析 ================= */
    function parseTable() {
        const table = document.getElementById("kbgrid_table_0");
        const divs = [];
        const week = [];

        if (table) {
            table.querySelectorAll("tr").forEach(row => {
                row.querySelectorAll("td[id]").forEach(td => {
                    const id = td.getAttribute("id");
                    if (!id) return;

                    const weekday = parseInt(id.split("-")[0], 10);
                    if (weekday < 1 || weekday > 7) return;

                    td.querySelectorAll("div").forEach(div => {
                        divs.push(div);
                        week.push(weekday);
                    });
                });
            });
        }
        return { divs, week };
    }

    function extractLabRows() {
        const tbody = document.querySelector('#sycjlrtabGrid tbody');
        if (!tbody) return [];
        return Array.from(tbody.querySelectorAll('tr.jqgrow'));
    }

    /* ================= 周次解析 ================= */
    function parseWeekText(text) {
        const result = [];
        const regex = /(\d+)(?:-(\d+))?周(单|双)?/g;
        let m;
        while ((m = regex.exec(text)) !== null) {
            result.push({
                start: +m[1],
                end: +(m[2] || m[1]),
                interval: m[3] ? 2 : 1
            });
        }
        return result;
    }

    /* ================= 提取校区+Bxx-xxx格式地点 ================= */
    function extractFormattedLocation(locationStr) {
        if (!locationStr) return null;

        const parts = locationStr.split('/');

        // 如果是理论课格式（只有一个部分），直接返回
        if (parts.length === 1) {
            return {
                formatted: parts[0],
                campus: parts[0].match(/.*校区/)?.[0] || "",
                room: parts[0].match(/B\d+-\d+/)?.[0] || "",
                labName: "",
                fullInfo: parts[0]
            };
        }

        // 实验课格式：实训室名称/Bxx-xxx/校区
        let campus = "";
        let room = "";
        let labName = "";

        // 提取校区（最后一部分）
        for (let i = parts.length - 1; i >= 0; i--) {
            if (parts[i].includes("校区")) {
                campus = parts[i];
                break;
            }
        }

        // 提取教室号（Bxx-xxx格式）
        for (let i = 0; i < parts.length; i++) {
            const roomMatch = parts[i].match(/B\d+-\d+/);
            if (roomMatch) {
                room = roomMatch[0];
                break;
            }
        }

        // 提取实训室名称（第一个部分，但不是校区和教室号）
        if (parts.length > 0 && !parts[0].includes("校区") && !parts[0].match(/B\d+-\d+/)) {
            labName = parts[0];
        }

        // 组合地点：校区+教室号（如果都有的话）
        let formattedLocation = "";
        if (campus && room) {
            formattedLocation = `${campus}${room}`;
        } else if (room) {
            formattedLocation = room;
        } else if (campus) {
            formattedLocation = campus;
        } else if (labName) {
            formattedLocation = labName;
        } else {
            formattedLocation = "实验室";
        }

        return {
            formatted: formattedLocation,
            campus: campus,
            room: room,
            labName: labName,
            fullInfo: locationStr
        };
    }

    /* ================= 课程解析 ================= */
    function parseCourses({ divs, week }) {
        const courses = [];

        /* ================= 普通课程 ================= */
        divs.forEach((div, i) => {
            const course = {
                name: extractCourseName(div),
                week: week[i],
                startWeek: [],
                endWeek: [],
                isSingleOrDouble: [],
                startTime: null,
                endTime: null,
                location: "",
                teacher: "",
                isLab: false,
                allLocations: [], // 保存所有地点的格式化字符串
                locationDetails: [] // 保存详细地点信息
            };

            div.querySelectorAll("p").forEach(p => {
                const span = p.querySelector("span[title]");
                if (!span) return;

                const key = span.title;
                const value = p.querySelectorAll("font")[1]?.innerText?.trim() || "";

                if (key.includes("节/周")) {
                    const t = value.match(/\((\d+)-(\d+)节\)/);
                    if (t) {
                        course.startTime = +t[1];
                        course.endTime = +t[2];
                    }

                    parseWeekText(value).forEach(w => {
                        course.startWeek.push(w.start);
                        course.endWeek.push(w.end);
                        course.isSingleOrDouble.push(w.interval);
                    });
                }

                if (key.includes("上课地点")) {
                    course.location = value;
                    // 为理论课也设置 allLocations 和 locationDetails
                    if (value) {
                        const locationInfo = extractFormattedLocation(value);
                        if (locationInfo) {
                            course.allLocations = [locationInfo.formatted];
                            course.locationDetails = [locationInfo];
                        }
                    }
                }
                if (key.includes("教师")) course.teacher = value;
            });

            if (course.startTime !== null) courses.push(course);
        });

        /* ================= 实验课 / 综合创新训练解析 ================= */
        const labRows = extractLabRows();

        if (labRows && labRows.length > 0) {
            // console.log("【实验课】找到实验课行数:", labRows.length);

            // 用于合并相同时间段但不同地点的课程
            const labCourseMap = new Map();

            labRows.forEach(row => {
                const cells = row.querySelectorAll("td[role='gridcell']");
                const data = {};
                cells.forEach(cell => {
                    const descId = cell.getAttribute("aria-describedby");
                    if (descId && descId.startsWith("sycjlrtabGrid_")) {
                        const field = descId.replace("sycjlrtabGrid_", "");
                        data[field] = cell.getAttribute('title') || cell.textContent.trim();
                    }
                });

                // console.log("【实验课】原始数据:", data);

                // 实验课时间地点字符串
                const timeLocation = String(data.sksjdd || "");
                const locationDetail = String(data.dycdxq || "");
                const projectName = String(data.xmmc || "");

                if (!timeLocation || timeLocation === "undefined" || timeLocation === "null" || !timeLocation.includes('[')) {
                    console.warn("【实验课】时间地点为空或格式错误，跳过课程:", data.kcmc);
                    return;
                }

                // console.log("【实验课】时间字符串:", timeLocation);
                // console.log("【实验课】地点字符串:", locationDetail);

                // 解析时间段
                const timeSegments = [];
                const rawSegments = timeLocation.split(',');

                rawSegments.forEach(s => {
                    const trimmed = s.trim();
                    if (trimmed && trimmed.includes('[') && trimmed.includes(']')) {
                        timeSegments.push(trimmed);
                    }
                });

                if (timeSegments.length === 0) {
                    console.warn("【实验课】未找到有效时间段，跳过课程:", data.kcmc);
                    return;
                }

                // console.log("【实验课】有效时间段:", timeSegments);

                // 解析地点（更健壮的分割方法）
                const locationSegments = [];
                if (locationDetail) {
                    // 尝试多种分割方式
                    const parts = locationDetail.split(/[,，]/); // 支持中文逗号和英文逗号
                    parts.forEach(part => {
                        const trimmed = part.trim();
                        if (trimmed) {
                            locationSegments.push(trimmed);
                        }
                    });
                }

                // console.log("【实验课】地点分割:", locationSegments);

                // 先收集所有时间段的地点信息
                const timeSegmentLocations = [];

                timeSegments.forEach((segment, index) => {
                    // console.log("【实验课】正在解析时间段:", segment);

                    // 获取对应的地点字符串
                    let locationFull = "";
                    if (locationSegments.length > index) {
                        locationFull = locationSegments[index];
                    } else if (locationSegments.length > 0) {
                        locationFull = locationSegments[0];
                    }

                    // 提取格式化地点信息
                    const locationInfo = extractFormattedLocation(locationFull);
                    if (locationInfo) {
                        timeSegmentLocations.push(locationInfo);
                    }
                });

                // console.log("【实验课】时间段对应地点:", timeSegmentLocations.map(loc => loc.formatted));

                // 解析第一个时间段的时间信息（假设所有时间段时间相同）
                const firstSegment = timeSegments[0];
                let weekdayChar = "";
                let startClass = 0;
                let endClass = 0;
                let startWeek = 0;
                let endWeek = 0;

                try {
                    // 提取星期
                    const weekdayMatch = firstSegment.match(/星期([一二三四五六日])/);
                    if (!weekdayMatch) {
                        console.error("【实验课】无法提取星期:", firstSegment);
                        return;
                    }
                    weekdayChar = weekdayMatch[1];

                    // 提取节次
                    const classMatch = firstSegment.match(/\[(\d+)-(\d+)节/);
                    if (!classMatch) {
                        console.error("【实验课】无法提取节次:", firstSegment);
                        return;
                    }
                    startClass = parseInt(classMatch[1], 10);
                    endClass = parseInt(classMatch[2], 10);

                    // 提取周次
                    const weekMatch = firstSegment.match(/(\d+)(?:-(\d+))?\s*周/);
                    if (!weekMatch) {
                        console.error("【实验课】无法提取周次:", firstSegment);
                        return;
                    }
                    startWeek = parseInt(weekMatch[1], 10);
                    endWeek = weekMatch[2] ? parseInt(weekMatch[2], 10) : startWeek;

                    // console.log("【实验课】解析结果:", {
                    //     星期: weekdayChar,
                    //     开始节次: startClass,
                    //     结束节次: endClass,
                    //     开始周: startWeek,
                    //     结束周: endWeek,
                    //     所有地点: timeSegmentLocations.map(loc => loc.formatted)
                    // });
                } catch (error) {
                    console.error("【实验课】解析时间段时出错:", error, firstSegment);
                    return;
                }

                // 创建唯一键，用于合并相同时间段的课程
                const courseKey = `${data.kcmc}_${startWeek}_${endWeek}_${weekdayChar}_${startClass}_${endClass}`;

                if (labCourseMap.has(courseKey)) {
                    // 已存在相同时间段的课程，合并地点信息
                    const existingCourse = labCourseMap.get(courseKey);
                    // 添加新地点到地点列表（去重）
                    timeSegmentLocations.forEach(locInfo => {
                        // 根据格式化后的地点字符串去重
                        const existingIndex = existingCourse.locationDetails.findIndex(
                            item => item.formatted === locInfo.formatted
                        );
                        if (existingIndex === -1) {
                            existingCourse.locationDetails.push(locInfo);
                        }
                    });

                    // 更新所有地点的格式化字符串（去重）
                    existingCourse.allLocations = existingCourse.locationDetails.map(loc => loc.formatted);
                    existingCourse.location = existingCourse.allLocations.join('、');

                    // 更新合并备注
                    if (existingCourse.locationDetails.length > 1) {
                        existingCourse.mergedNote = `多个实验地点：${existingCourse.locationDetails.map(loc =>
                            loc.labName ? `${loc.formatted}（${loc.labName}）` : loc.formatted
                        ).join('、')}`;
                    } else if (existingCourse.locationDetails.length === 1 && existingCourse.locationDetails[0].labName) {
                        existingCourse.mergedNote = `实验地点：${existingCourse.locationDetails[0].formatted}（${existingCourse.locationDetails[0].labName}）`;
                    }

                    // console.log("【实验课】合并地点到现有课程:", courseKey, "地点列表:", existingCourse.allLocations);
                } else {
                    // 创建新课程
                    const allFormattedLocations = timeSegmentLocations.map(loc => loc.formatted);
                    const uniqueLocations = [...new Set(allFormattedLocations)];

                    const course = {
                        name: String(data.kcmc || "实验课"),
                        isLab: true,
                        detail: projectName,
                        week: weekDayMap[weekdayChar] || 1,
                        startTime: startClass,
                        endTime: endClass,
                        startWeek: [startWeek],
                        endWeek: [endWeek],
                        isSingleOrDouble: [1],
                        // 多个地点用顿号连接
                        location: uniqueLocations.join('、'),
                        teacher: String(data.jsxm || ""),
                        allLocations: uniqueLocations,
                        locationDetails: timeSegmentLocations,
                        mergedNote: ""
                    };

                    // 设置备注
                    if (timeSegmentLocations.length > 1) {
                        course.mergedNote = `多个实验地点：${timeSegmentLocations.map(loc =>
                            loc.labName ? `${loc.formatted}（${loc.labName}）` : loc.formatted
                        ).join('、')}`;
                    } else if (timeSegmentLocations.length === 1 && timeSegmentLocations[0].labName) {
                        course.mergedNote = `实验地点：${timeSegmentLocations[0].formatted}（${timeSegmentLocations[0].labName}）`;
                    }

                    labCourseMap.set(courseKey, course);
                    // console.log("【实验课】已添加新课程:", {
                    //     name: course.name,
                    //     周次: `${startWeek}-${endWeek}周`,
                    //     星期: course.week,
                    //     节次: `${startClass}-${endClass}节`,
                    //     地点: course.location,
                    //     所有地点: course.allLocations,
                    //     地点详情: course.locationDetails,
                    //     合并备注: course.mergedNote
                    // });
                }
            });

            // 将合并后的实验课添加到courses数组
            labCourseMap.forEach(course => {
                // 确保地点信息正确
                if (course.locationDetails.length > 1 && !course.mergedNote) {
                    course.mergedNote = `多个实验地点：${course.locationDetails.map(loc =>
                        loc.labName ? `${loc.formatted}（${loc.labName}）` : loc.formatted
                    ).join('、')}`;
                } else if (course.locationDetails.length === 1 && course.locationDetails[0].labName && !course.mergedNote) {
                    course.mergedNote = `实验地点：${course.locationDetails[0].formatted}（${course.locationDetails[0].labName}）`;
                }
                courses.push(course);
                // console.log("【实验课】最终添加课程:", {
                //     name: course.name,
                //     周次: `${course.startWeek[0]}-${course.endWeek[0]}周`,
                //     星期: course.week,
                //     节次: `${course.startTime}-${course.endTime}节`,
                //     地点: course.location,
                //     所有地点: course.allLocations,
                //     地点详情: course.locationDetails,
                //     合并备注: course.mergedNote
                // });
            });
        }

        console.log("【ICS】课程总数:", courses.length);
        return courses;
    }

    /* ================= 其余工具函数 ================= */

    function extractCourseName(div) {
        return div.querySelector(".title font")?.innerText.trim()
            || div.querySelector(".title")?.innerText.trim()
            || "未知课程";
    }

    const TIME = {
        1: [8, 0], 2: [8, 55], 3: [10, 0], 4: [10, 55],
        5: [14, 0], 6: [14, 55], 7: [16, 0], 8: [16, 55],
        9: [19, 0], 10: [19, 55]
    };

    const pad = n => (n < 10 ? "0" : "") + n;

    function formatDT(d) {
        return d.getFullYear()
            + pad(d.getMonth() + 1)
            + pad(d.getDate())
            + "T"
            + pad(d.getHours())
            + pad(d.getMinutes())
            + "00";
    }

    function formatDTUTC(d) {
        return d.getUTCFullYear()
            + pad(d.getUTCMonth() + 1)
            + pad(d.getUTCDate())
            + "T"
            + pad(d.getUTCHours())
            + pad(d.getUTCMinutes())
            + "00";
    }

    function getDate(startDate, week, weekday) {
        const d = new Date(startDate);
        d.setDate(d.getDate() + (week - 1) * 7 + (weekday - 1));
        return d;
    }

    function generateUID(course, week) {
        const raw = [course.name, course.week, course.startTime, course.endTime, week].join("|");
        let hash = 0;
        for (let i = 0; i < raw.length; i++) {
            hash = ((hash << 5) - hash) + raw.charCodeAt(i);
            hash |= 0;
        }
        return `gdust-${Math.abs(hash)}@jwpt`;
    }

    function generateFileName(startDate) {
        const d = new Date(startDate);
        const y = d.getFullYear();
        const m = d.getMonth() + 1;
        return m >= 8 ? `${y}-${y + 1}-第1学期-课表.ics`
                      : `${y - 1}-${y}-第2学期-课表.ics`;
    }

    function generateCalendar(courses, startDate) {
        const lines = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//GDUST//Class Schedule//CN",
            "CALSCALE:GREGORIAN"
        ];

        courses.forEach(c => {
            for (let i = 0; i < c.startWeek.length; i++) {
                for (let wk = c.startWeek[i]; wk <= c.endWeek[i]; wk++) {
                    if (c.isSingleOrDouble[i] === 2 && (wk - c.startWeek[i]) % 2) continue;

                    const base = getDate(startDate, wk, c.week);
                    const start = new Date(base);
                    const end = new Date(base);

                    start.setHours(...TIME[c.startTime]);
                    end.setHours(TIME[c.endTime][0], TIME[c.endTime][1] + 45);

                    // console.log(
                    //     "【DEBUG】生成事件",
                    //     c.name,
                    //     "第", wk, "周",
                    //     "星期", c.week,
                    //     "节次", c.startTime, "-", c.endTime,
                    //     "地点:", c.location,
                    //     "所有地点:", c.allLocations || []
                    // );

                    pushEvent(lines, c, wk, start, end);
                }
            }
        });

        lines.push("END:VCALENDAR");

        const blob = new Blob([lines.join(CRLF)], { type: "text/calendar;charset=utf-8" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = generateFileName(startDate);
        a.click();
    }

    function pushEvent(lines, course, week, start, end) {
        lines.push("BEGIN:VEVENT");
        lines.push(`UID:${generateUID(course, week)}`);
        lines.push(`DTSTART;TZID=${TZID}:${formatDT(start)}`);
        lines.push(`DTEND;TZID=${TZID}:${formatDT(end)}`);
        lines.push(`DTSTART:${formatDTUTC(start)}Z`);
        lines.push(`DTEND:${formatDTUTC(end)}Z`);
        lines.push(`SUMMARY:${course.isLab ? "实验课-" : ""}${course.name}`);

        // 设置日程地点（可能包含多个地点，用顿号分隔）
        if (course.location) {
            lines.push(`LOCATION:${course.location}`);
        }

        const descParts = [];
        descParts.push(`第${week}周`);
        if (course.isLab && course.detail) {
            descParts.push(`实验项目：${course.detail}`);
        }
        if (course.teacher) {
            descParts.push(`教师：${course.teacher}`);
        }

        // 添加地点备注（实验课需要详细信息）
        if (course.isLab) {
            if (course.mergedNote) {
                descParts.push(course.mergedNote);
            } else if (course.locationDetails && course.locationDetails.length > 0) {
                if (course.locationDetails.length > 1) {
                    descParts.push(`多个实验地点：${course.locationDetails.map(loc =>
                        loc.labName ? `${loc.formatted}（${loc.labName}）` : loc.formatted
                    ).join('、')}`);
                } else if (course.locationDetails[0].labName) {
                    descParts.push(`实验地点：${course.locationDetails[0].formatted}（${course.locationDetails[0].labName}）`);
                }
            }
        }

        if (descParts.length > 0) {
            lines.push(`DESCRIPTION:${descParts.join("\\n")}`);
        }

        lines.push("END:VEVENT");
    }
})();