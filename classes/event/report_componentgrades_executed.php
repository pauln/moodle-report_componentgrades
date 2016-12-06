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
 * report_componentgrades_executed
 *
 * @package    report_componentgrades
 * @author     Suan Kan <suankan@catalyst-au.net>
 * @copyright  2016 Catalyst IT
 */
namespace report_componentgrades\event;
defined('MOODLE_INTERNAL') || die();
/**
 * report_componentgrades_executed
 *
 * This event is to be triggered whenever report_componentgrades is executed to export data.
 *
 **/
class report_componentgrades_executed extends \core\event\base {
    protected function init() {
        $this->data['crud'] = 'u';
        $this->data['edulevel'] = self::LEVEL_OTHER;
        $this->context = \context_system::instance();
    }

    public static function get_name() {
        return get_string('eventreportcomponentgradesexecuted', 'report_componentgrades');
    }

    public function get_description() {
        $coursefullname = $this->data['other']['coursefullname'];
        $courseid = $this->data['other']['courseid'];
        $assignmentname = $this->data['other']['assignmentname'];

        $desc = "Export of report_componentgrades executed for assignment $assignmentname in course $courseid, $coursefullname";

        return $desc;
    }

}
