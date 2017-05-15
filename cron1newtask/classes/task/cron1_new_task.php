<?php
namespace local_cron1newtask\task;
class cron1_new_task extends \core\task\scheduled_task {
    /**
     * Get a descriptive name for this task (shown to admins).
     *
     * @return string
     */
    public function get_name() {
        return get_string('pluginname', 'local_cron1newtask');
    }

    /**
     * Run update_status cron.
     */
    public function execute() {
        mtrace('Mail send for upcoming event2');
        // Update ticket status
        $this->mail_send_upcoming_event2();
    }
    function mail_send_upcoming_event2() {
        global $DB,$USER, $SESSION, $CFG;
        ini_set('max_execution_time',0);
        $starttime = microtime();
        define('FULLME', 'cron');
        $nomoodlecookie = true;

        if (!isset($_SERVER['REMOTE_ADDR']) && isset($_SERVER['argv'][0])) {
        chdir(dirname($_SERVER['argv'][0]));
        }
        //require_once(dirname(__FILE__) . '../../../config.php');
        require_once($CFG->libdir.'/adminlib.php');
        require_once($CFG->libdir.'/gradelib.php');

        if (!empty($CFG->showcronsql)) {
        $DB->debug = true;
        }

        if (!empty($CFG->showcrondebugging)) {
        $CFG->debug = DEBUG_DEVELOPER;
        $CFG->debugdisplay = true;
        }
        // extra safety

        @session_write_close();
        // check if execution allowed

        if (isset($_SERVER['REMOTE_ADDR'])) { // if the script is accessed via the web.
            if (!empty($CFG->cronclionly)) {
                // This script can only be run via the cli.
                print_error(get_string('cronerror', 'local_cron1newtask'), 'admin');
                exit;
            }
        // This script is being called via the web, so check the password if there is one.

            if (!empty($CFG->cronremotepassword)) {
                $pass = optional_param('password', '', PARAM_RAW);
                if($pass != $CFG->cronremotepassword) {
                // wrong password.
                    print_error(get_string('cronpswd', 'local_cron1newtask'), 'admin');
                    exit;
                }
            }
        }
        // emulate normal session
        //these two line are not standared in moodle 3.2 so we are going to commenting line
        //$SESSION = new object();
        //$USER = get_admin();      /// Temporarily, to provide environment for this script
        // ignore admins timezone, language and locale - use site deafult instead!

        $USER->timezone = $CFG->timezone;
        $USER->lang = '';
        $USER->theme = '';
        //course_setup(SITEID);
        // send mime type and encoding

       // if (check_browser_version('MSIE')) {
        //ugly IE hack to work around downloading instead of viewing

         //   @header('Content-Type: text/html; charset=utf-8');
         //   echo "<xmp>"; //<pre> is not good enough for us here
       // } else {
            //send proper plaintext header
         //   @header('Content-Type: text/plain; charset=utf-8');
       // }
        // no more headers and buffers

        while(@ob_end_flush());
        // increase memory limit (PHP 5.2 does different calculation, we need more memory now)

        @ini_set('memory_limit','1000M');
        // Start output log
        $timenow  = time();
        mtrace("Server Time: ".date('r',$timenow)."\n\n");
        mtrace('Starting main gradebook job ...');

        //$courseUserRS = get_recordset_sql($reminderUserID);
        ////Naga added to send the details in excel sheet
        $xlshead=pack("s*", 0x809, 0x8, 0x0, 0x10, 0x0, 0x0);
        $xlsfoot=pack("s*", 0x0A, 0x00);
        function xlsCell($row,$col,$val) {
            $len=strlen($val);
            return pack("s*",0x204,8+$len,$row,$col,0x0,$len).$val;
        }
        $data = '';
        $data.=xlsCell(0,0,"Name") . xlsCell(0,1,"Portal ID") . xlsCell(0,2,"Course Name") .
            xlsCell(0,3,"Course Completion date") . xlsCell(0,4,"Status") . xlsCell(0,5,"Question1") .
            xlsCell(0,6,"Response1"). xlsCell(0,7,"Question2") . xlsCell(0,8,"Response2"). xlsCell(0,9,"Question3") .
            xlsCell(0,10,"Response3"). xlsCell(0,11,"Question4") . xlsCell(0,12,"Response4"). xlsCell(0,13,"Question5") .
            xlsCell(0,14,"Response5"). xlsCell(0,15,"Question6") . xlsCell(0,16,"Response6"). xlsCell(0,17,"Question7") .
            xlsCell(0,18,"Response7"). xlsCell(0,19,"Question8") . xlsCell(0,20,"Response8"). xlsCell(0,21,"Question9") .
            xlsCell(0,22,"Response9") . xlsCell(0,23,"Question10") . xlsCell(0,24,"Response10");
        $rowNumber=0;
//$query = "SELECT t.id as id ,t.value as value, u.username as portalid FROM mdl_scorm_scoes_track t join mdl_user u on u.id=t.userid where t.scormid=772 and t.element='cmi.core.lesson_status'";
//$query = "select t.value,s.name,u.username,concat(u.firstname,' ',u.lastname)as name from mdl_scorm_scoes_track t join mdl_scorm s on s.id=t.scormid join mdl_user u on u.id=t.userid where t.scormid=1306 and t.element='cmi.interactions_0.student_response'";
//$querycheck="select distinct(userid) from mdl_scorm_scoes_track where scormid=2634 ";

        $querycheck="select distinct(userid)
            from {scorm_scoes_track}
            where scormid=2634 and userid not in (select distinct(userid)
            from {scorm_scoes_track} where value='SENT TO LEGAL' and scormid=2634)";
        if ($courseUserRS12 = $DB->get_records_sql($querycheck)) {
           foreach($courseUserRS12 as $UserId1) {
                $query=" select concat(u.firstname,' ',u.lastname)as name,u.username as portalid,s.name as coursename,
                    from_unixtime(gg.timemodified) as completiondate,
                    (select k.value  from {scorm_scoes_track} k join {user} u on u.id=k.userid
                    where k.element = ('cmi.core.lesson_status') and u.id = $UserId1->userid and k.scormid=2634 ) as status,
                    (select a.value from {scorm_scoes_track} a join {user} u
                    on u.id=a.userid where a.element = ('cmi.interactions_0.student_response')
                    and u.id=$UserId1->userid and a.scormid=2634 ) as response1,
                    (select b.value  from {scorm_scoes_track} b  join {user} u
                    on u.id=b.userid where b.element = ('cmi.interactions_1.student_response')
                    and u.id=$UserId1->userid and b.scormid=2634 ) as response2,
                    (select c.value  from {scorm_scoes_track} c join {user} u
                    on u.id=c.userid where c.element = ('cmi.interactions_2.student_response')
                    and u.id=$UserId1->userid and c.scormid=2634 ) as response3,
                    (select d.value  from {scorm_scoes_track} d  join {user} u
                    on u.id=d.userid where d.element = ('cmi.interactions_3.student_response')
                    and u.id=$UserId1->userid and d.scormid=2634 ) as response4,
                    (select e.value  from {scorm_scoes_track} e join {user} u
                    on u.id=e.userid where e.element = ('cmi.interactions_4.student_response')
                    and u.id=$UserId1->userid and e.scormid=2634 ) as response5,
                    (select f.value  from {scorm_scoes_track} f  join {user} u
                    on u.id=f.userid where f.element = ('cmi.interactions_5.student_response')
                    and u.id=$UserId1->userid and f.scormid=2634 ) as response6,
                    (select g.value  from {scorm_scoes_track} g join {user} u
                    on u.id=g.userid where g.element = ('cmi.interactions_6.student_response')
                    and u.id=$UserId1->userid and g.scormid=2634 ) as response7,
                    (select h.value  from {scorm_scoes_track} h join {user} u
                    on u.id=h.userid where h.element = ('cmi.interactions_7.student_response')
                    and u.id=$UserId1->userid and h.scormid=2634 ) as response8,
                    (select i.value  from {scorm_scoes_track} i join {user} u
                    on u.id=i.userid  where i.element = ('cmi.interactions_8.student_response')
                    and u.id=$UserId1->userid and i.scormid=2634 ) as response9,
                    (select j.value  from {scorm_scoes_track} j join {user} u
                    on u.id=j.userid  where j.element = ('cmi.interactions_9.student_response')
                    and u.id=$UserId1->userid and j.scormid=2634 ) as response10
                    from {scorm_scoes_track} t  join {scorm} s on s.id=t.scormid join {user} u
                    on u.id=t.userid
                    JOIN {grade_items} AS gi on gi.itemname=s.name join {grade_grades} AS gg
                    on gg.userid= t.userid and gi.id = gg.itemid and gg.finalgrade >= gi.grademax
                    where t.scormid=2634 and gi.iteminstance=2634 and t.element like'%student_response%'
                    and u.id=$UserId1->userid group by u.id";
            $query1=$DB->get_records_sql($query);
                foreach($query1 as $books) {
                    $rowNumber=$rowNumber+1;
                    $name=$books->name;
                    $portalid=$books->portalid;
                    $coursename=$books->coursename;
                    $completiondate=$books->completiondate;
                    $status=$books->status;
                    if($books->response10=='') {
                        $response1='';
                        $response2=$books->response1;
                        $response3=$books->response2;
                        $response4=$books->response3;
                        $response5=$books->response4;
                        $response6=$books->response5;
                        $response7=$books->response6;
                        $response8=$books->response7;
                        $response9=$books->response8;
                        $response10=$books->response9;
                    } else{
                        $response1=$books->response1;
                        $response2=$books->response2;
                        $response3=$books->response3;
                        $response4=$books->response4;
                        $response5=$books->response5;
                        $response6=$books->response6;
                        $response7=$books->response7;
                        $response8=$books->response8;
                        $response9=$books->response9;
                        $response10=$books->response10;
                    }
                    $Question1 = get_string('question1', 'local_cron1newtask');
                    $Question2 = get_string('question2', 'local_cron1newtask');
                    $Question3 = get_string('question3', 'local_cron1newtask');
                    $Question4 = get_string('question4', 'local_cron1newtask');
                    $Question5 = get_string('question5', 'local_cron1newtask');
                    $Question6 = get_string('question6', 'local_cron1newtask');
                    $Question7 = get_string('question7', 'local_cron1newtask');
                    $Question8 = get_string('question8', 'local_cron1newtask');
                    $Question9 = get_string('question9', 'local_cron1newtask');
                    $Question10= get_string('question10','local_cron1newtask');
                    $data.=xlsCell($rowNumber,0,$name) . xlsCell($rowNumber,1,$portalid) . xlsCell($rowNumber,2,$coursename) .
        xlsCell($rowNumber,3,$completiondate). xlsCell($rowNumber,4,$status) . xlsCell($rowNumber,5,$Question1) .
        xlsCell($rowNumber,6,$response1). xlsCell($rowNumber,7,$Question2). xlsCell($rowNumber,8,$response2).
        xlsCell($rowNumber,9,$Question3). xlsCell($rowNumber,10,$response3). xlsCell($rowNumber,11,$Question4).
        xlsCell($rowNumber,12,$response4). xlsCell($rowNumber,13,$Question5). xlsCell($rowNumber,14,$response5).
        xlsCell($rowNumber,15,$Question6). xlsCell($rowNumber,16,$response6). xlsCell($rowNumber,17,$Question7).
        xlsCell($rowNumber,18,$response7). xlsCell($rowNumber,19,$Question8). xlsCell($rowNumber,20,$response8).
        xlsCell($rowNumber,21,$Question9). xlsCell($rowNumber,22,$response9). xlsCell($rowNumber,23,$Question10).
        xlsCell($rowNumber,24,$response10);
                }
            }
        }
        $uniqueid=  substr(md5(uniqid(rand(), true)),16,16);
        $filen="CoBC LC Responses";
        $filename="$filen.xls";
        $conte=$xlshead . $data . $xlsfoot;//This is previous not there please place it
        file_put_contents('/var/www/html/ntt/local/cron1newtask/'.$filename, $conte);//This is the command to save your file in server
        $content = chunk_split(base64_encode($conte));
        $uid = md5(uniqid(time()));
        $agname="Name Of Recipent";
        $message1 ='<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:w="urn:schemas-microsoft-com:office:word"
            xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
            xmlns="http://www.w3.org/TR/REC-html40">
            <head>
                <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
                <meta name=Generator content="Microsoft Word 14 (filtered medium)">
                <style>
                    <!--/* Font Definitions */
                        @font-face
                        {
                            font-family:"Microsoft Sans Serif";
                            panose-1:2 11 6 4 2 2 2 2 2 4;
                        }
                        @font-face
                        {
                            font-family:Calibri;
                            panose-1:2 15 5 2 2 2 4 3 2 4;
                        }
                        /* Style Definitions */
                        p.MsoNormal, li.MsoNormal, div.MsoNormal
                        {
                            margin-top:0in;
                            margin-right:0in;
                            margin-bottom:10.0pt;
                            margin-left:0in;
                            line-height:115%;
                            font-size:11.0pt;
                             font-family:"Calibri","sans-serif";
                        }
                        a:link, span.MsoHyperlink
                        {
                            mso-style-priority:99;
                            color:blue;
                            text-decoration:underline;
                        }
                        a:visited, span.MsoHyperlinkFollowed
                        {
                            mso-style-priority:99;
                            color:purple;
                            text-decoration:underline;
                        }
                        span.EmailStyle17
                        {
                            mso-style-type:personal-compose;
                            font-family:"Calibri","sans-serif";
                            color:windowtext;
                        }
                        .MsoChpDefault
                        {
                            mso-style-type:export-only;
                            font-family:"Calibri","sans-serif";
                        }
                        @page WordSection1
                        {
                            size:8.5in 11.0in;
                            margin:1.0in 1.0in 1.0in 1.0in;
                        }
                        div.WordSection1
                        {
                            page:WordSection1;
                        }
                        --></style><!--[if gte mso 9]><xml>
                <o:shapedefaults v:ext="edit" spidmax="1026" />
                </xml><![endif]--><!--[if gte mso 9]><xml>
                <o:shapelayout v:ext="edit">
                <o:idmap v:ext="edit" data="1" />
                </o:shapelayout></xml><![endif]--></head><body lang=EN-US link=blue vlink=purple>
                <div class=WordSection1>
                    <p class=MsoNormal >Hi Katrina,<o:p></o:p></p>
                    <p class=MsoNormal>This message is generated from the CATALYS system capturing Code of Business Conduct
                        input from the Leadership Council versions of the course.Attached are the reports for the current week.
                        Please acknowledge receipt of the attached and confirm that you have the records presented here.
                        Your tracking system will be the system of record and the entries will be replaced in CATALYS with
                        the message accordingly.<o:p></o:p>
                    </p>
                    <p class=MsoNormal>Should you have any questions about this data feed,
                        please contact Janet Lahlou, Talent Development.&nbsp;
                    <p class=MsoNormal>Thank you,<o:p></o:p>
                    </p>
                    <p class=MsoNormal>CATALYS System Team<o:p></o:p></p></div></body></html>';

         $message2 ='
            <html xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:w="urn:schemas-microsoft-com:office:word"
                xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
                xmlns="http://www.w3.org/TR/REC-html40">
                <head>
                    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
                    <meta name=Generator content="Microsoft Word 14 (filtered medium)">
                    <style><!--/* Font Definitions */
                    @font-face
                    {
                        font-family:"Microsoft Sans Serif";
                        panose-1:2 11 6 4 2 2 2 2 2 4;
                    }
                    @font-face
                    {
                        font-family:Calibri;
                        panose-1:2 15 5 2 2 2 4 3 2 4;
                    }
                    /* Style Definitions */
                    p.MsoNormal, li.MsoNormal, div.MsoNormal
                    {
                        margin-top:0in;
                        margin-right:0in;
                        margin-bottom:10.0pt;
                        margin-left:0in;
                        line-height:115%;
                        font-size:11.0pt;
                        font-family:"Calibri","sans-serif";
                    }
                    a:link, span.MsoHyperlink
                    {
                        mso-style-priority:99;
                        color:blue;
                        text-decoration:underline;
                    }
                    a:visited, span.MsoHyperlinkFollowed
                    {
                        mso-style-priority:99;
                        color:purple;
                        text-decoration:underline;
                    }
                    span.EmailStyle17
                    {
                        mso-style-type:personal-compose;
                        font-family:"Calibri","sans-serif";
                        color:windowtext;
                    }
                    .MsoChpDefault
                    {
                        mso-style-type:export-only;
                        font-family:"Calibri","sans-serif";
                    }
                    @page WordSection1
                    {
                        size:8.5in 11.0in;
                        margin:1.0in 1.0in 1.0in 1.0in;
                    }
                    div.WordSection1
                    {
                        page:WordSection1;
                    }
                    --></style>
                    <!--[if gte mso 9]><xml>
                        <o:shapedefaults v:ext="edit" spidmax="1026" />
                        </xml><![endif]--><!--[if gte mso 9]><xml>
                        <o:shapelayout v:ext="edit">
                        <o:idmap v:ext="edit" data="1" />
                        </o:shapelayout></xml><![endif]--></head><body lang=EN-US link=blue vlink=purple>
                    <div class=WordSection1>
                    <p class=MsoNormal >Hi Katrina,<o:p></o:p></p>
                        <p class=MsoNormal>This message is generated from the CATALYS system capturing Code of
                            Business Conduct input from the Leadership Council version of the course of the course.
                            There are no new records for the LC level Course.
                            <o:p></o:p>
                        </p>
                    <p class=MsoNormal>Should you have any questions about this data feed, please contact Janet Lahlou,
                        Talent Development.&nbsp;
                    <p class=MsoNormal>Thank you,<o:p></o:p></p>
                        <p class=MsoNormal>CATALYS System Team<o:p></o:p></p>
                    </div>
                    </body>
            </html>';
            //$header = "From: Place Website Name <".$from_mail.">\r\n";
            // $header .= "Reply-To: ".$from_mail."\r\n";
            //$header .= "MIME-Version: 1.0\r\n";
            //$header="CoBC 2015 Responses for LC Group";
            //
            $header = '';
            $header .= "CoBC 2015 Responses for LC Group\r\n";
            $header .= "Content-Type: multipart/mixed; boundary=\"".$uid."\"\r\n\r\n";
            $header .= "This is a multi-part message in MIME format.\r\n";
            $header .= "--".$uid."\r\n";
            $header .= "Content-type:text/html; charset=iso-8859-1\r\n";
            $header .= "Content-Transfer-Encoding: 7bit\r\n\r\n";
            $header .= $message1."\r\n\r\n";
            $header .= "--".$uid."\r\n";
            $header .= "Content-Type: application/vnd.ms-excel; name=\"".$filename."\"\r\n";
            $header .= "Content-Transfer-Encoding: base64\r\n";
            $header .= "Content-Disposition: attachment; filename=\"".$filename."\"\r\n\r\n";
            $header .= $content."\r\n\r\n";
            $header .= "--".$uid."--";
            $tmpfile = '/COBC/CoBC LC Responses.xls';
            $file='CoBC LC Responses.xls';
            $reminderUserID = "SELECT * from {user} where username in ('admin2')";
            if ($courseUserRS =$DB->get_records_sql($reminderUserID)) {
                foreach($courseUserRS as $UserId) {
                    mtrace($UserId->id);
                    $euser = $DB->get_record('user',array('id' => $UserId->id));
                    if(!empty($querycheck)) {
                        email_to_user( $euser,'','COBC Compliance Team Report from LMS',$message1,$message1,$tmpfile,$file);
                    } else{
                        email_to_user( $euser,'','COBC Compliance Team Report from LMS',$message2,$message2);
                    }
                    //email_to_user( $euser,$header,'CoBC 2015 Responses for LC Group',$message1);
        //email_to_user( $euser,'CoBC 2015 training program Responses',$message1);
                }
            }
            mtrace('done.');
            //Unset session variables and destroy it
            @session_unset();
            @session_destroy();
            mtrace("Cron script completed correctly");
            $difftime = microtime_diff($starttime, microtime());
            mtrace("Execution took ".$difftime." seconds");
            // finish the IE hack
            //if (check_browser_version('MSIE')) {
            //    echo "</xmp>";
            //}
    }
}


