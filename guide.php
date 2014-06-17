<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Exports an Excel spreadsheet of the component grades in a Marking Guide-graded assignment.
 *
 * @package    report_componentgrades
 * @copyright  2014 Paul Nicholls
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

require('../../config.php');
require_once($CFG->dirroot.'/lib/excellib.class.php');

$id          = required_param('id', PARAM_INT);// Course ID
$modid       = required_param('modid', PARAM_INT);// CM ID

$params['id'] = $id;
$params['modid'] = $id;

$PAGE->set_url('/report/componentgrades/index.php', $params);

$course = $DB->get_record('course', array('id'=>$id), '*', MUST_EXIST);
require_login($course);

$modinfo = get_fast_modinfo($course->id);
$cm = $modinfo->get_cm($modid);
$modcontext = context_module::instance($cm->id);
require_capability('mod/assign:grade', $modcontext);

add_to_log($course->id, "course", "report componentgrades", "report/componentgrades/guide.php?id=$course->id&modid=$cm->id", $cm->id);

$filename = $course->shortname . ' - ' . $cm->name . '.xls';

$data = $DB->get_records_sql('SELECT ggf.id AS ggfid, crs.shortname AS course, asg.name AS assignment, gd.name AS guide, ggc.shortname, ggf.score, ggf.remark, ggf.criterionid, rubm.username AS grader, stu.id AS userid, stu.idnumber AS idnumber, stu.firstname, stu.lastname, stu.username AS student, FROM_UNIXTIME(gin.timemodified) AS modified
FROM {course} AS crs
JOIN {course_modules} AS cm ON crs.id = cm.course
JOIN {assign} AS asg ON asg.id = cm.instance
JOIN {context} AS c ON cm.id = c.instanceid
JOIN {grading_areas} AS ga ON c.id=ga.contextid
JOIN {grading_definitions} AS gd ON ga.id = gd.areaid
JOIN {gradingform_guide_criteria} AS ggc ON (ggc.definitionid = gd.id)
JOIN {grading_instances} AS gin ON gin.definitionid = gd.id
JOIN {assign_grades} AS ag ON ag.id = gin.itemid
JOIN {user} AS stu ON stu.id = ag.userid
JOIN {user} AS rubm ON rubm.id = gin.raterid
JOIN {gradingform_guide_fillings} AS ggf ON (ggf.instanceid = gin.id)
AND (ggf.criterionid = ggc.id)
WHERE cm.id = ? AND gin.status = 1
ORDER BY lastname ASC, firstname ASC, userid ASC, ggc.sortorder ASC, ggc.shortname ASC', array($cm->id));

$students = $DB->get_records_sql('SELECT stu.id AS userid, stu.idnumber AS idnumber, stu.firstname, stu.lastname, stu.username AS student
    FROM {user} AS stu
    JOIN {user_enrolments} ue ON ue.userid = stu.id
    JOIN {enrol} enr ON ue.enrolid = enr.id
    WHERE enr.courseid = ?
    ORDER BY lastname ASC, firstname ASC, userid ASC', array($course->id));

$first = reset($data);
if ($first === false) {
    $url = $CFG->wwwroot.'/mod/assign/view.php?id='.$cm->id;
    $message = "No grades have been entered into this assignment's marking guide.";
    redirect($url, $message, 5);
    exit;
}

$workbook = new MoodleExcelWorkbook("-");
$workbook->send($filename);
$sheet = $workbook->add_worksheet($cm->name);

// Course / assignment / marking guide names.
$sheet->write_string(0, 0, $course->fullname);
$format = $workbook->add_format(array('size'=>18, 'bold'=>1));
$sheet->set_row(0, 24, $format);
$sheet->write_string(1, 0, $cm->name);
$format = $workbook->add_format(array('size'=>16, 'bold'=>1));
$sheet->set_row(1, 21, $format);
$sheet->write_string(2, 0, 'Marking guide: ' . $first->guide);
$sheet->set_row(2, 21, $format);

// Column headers - two rows for grouping.
$sheet->write_string(4, 0, 'Student');
$sheet->merge_cells(4, 0, 4, 3);
$sheet->write_string(5, 0, 'First name');
$sheet->write_string(5, 1, 'Last name');
$sheet->write_string(5, 2, 'Student ID');
$sheet->write_string(5, 3, 'Username');
$sheet->set_column(0, 3, 10); // Set column widths to 10.

$pos = 4;
foreach($data as $line) {
    if ($line->userid !== $first->userid) {
        break;
    }
    $sheet->write_string(4, $pos, $line->shortname);
    $sheet->merge_cells(4, $pos, 4, $pos+1);
    $sheet->write_string(5, $pos, 'Score');
    $sheet->set_column($pos, $pos++, 6); // Set column width to 6.
    $sheet->write_string(5, $pos, 'Feedback');
    $sheet->set_column($pos, $pos++, 10); // Set column widths to 10.
}

$gradinginfopos = $pos;
$sheet->write_string(4, $pos, 'Grading info');
$sheet->write_string(5, $pos, 'Graded by');
$sheet->set_column($pos, $pos++, 10); // Set column width to 10.
$sheet->write_string(5, $pos, 'Time graded');
$sheet->set_column($pos, $pos, 17.5); // Set column width to 17.5.
$sheet->merge_cells(4, $pos-1, 4, $pos);

$format = $workbook->add_format(array('size'=>12, 'bold'=>1));
$sheet->set_row(4, 15, $format);
$format = $workbook->add_format(array('bold'=>1));
$sheet->set_row(5, null, $format);

// Merge header cells.
$sheet->merge_cells(0, 0, 0, $pos);
$sheet->merge_cells(1, 0, 1, $pos);
$sheet->merge_cells(2, 0, 2, $pos);

// Actual data.
$lastuser = 0;
$row = 6;
$datarow = 0;
$keys = array_keys($data);
$key = $keys[0];
$line = $data[$key];
foreach ($students as $student) {
    $row++;
    $sheet->write_string($row, 0, $student->firstname);
    $sheet->write_string($row, 1, $student->lastname);
    $sheet->write_string($row, 2, $student->idnumber);
    $sheet->write_string($row, 3, $student->student);

    $pos = 4;
    while($line->userid == $student->userid) {
        $sheet->write_number($row, $pos++, $line->score);
        $sheet->write_string($row, $pos++, $line->remark);
        if ($pos === $gradinginfopos) {
            $sheet->write_string($row, $pos++, $line->grader);
            $sheet->write_string($row, $pos, $line->modified);
        }
        $datarow++;
        $key = $keys[$datarow];
        $line = $data[$key];
    }
}

$workbook->close();

exit;