<?php
ini_set('memory_limit','-1');
class rbz_report extends CI_Controller {

	function __construct() {
		parent::__construct();
		$this->load->library('Excel');
		date_default_timezone_set('Africa/johannesburg');

	}

	function index() {
		$post = $this->input->post('report_obj');
		$data = json_decode($post);
		$year_date = new DateTime($data->year_date);
		$start_date = new DateTime($data->start_date);
		$end_date = new DateTime($data->end_date);
		$year_date = $year_date->format('Y-m-d');
		$start_date = $start_date->format('Y-m-d');
		$end_date = $end_date->format('Y-m-d');
		$after_period_date = strtotime($end_date.'+1 days');
		$after_period_date = date('Y-m-d',$after_period_date);
		$twelve_months_date = strtotime($end_date.'+12 months');
		$twelve_months_date = date('Y-m-d',$twelve_months_date);
		$thirty_date = strtotime($end_date.'-29 days');
		$thirty_date = date('Y-m-d',$thirty_date);
		$sixty_date = strtotime($end_date.'-59 days');
		$sixty_date = date('Y-m-d',$sixty_date);
		$ninety_date = strtotime($end_date.'-90 days');
		$ninety_date = date('Y-m-d',$ninety_date);
		$ninetyplus_date = strtotime($end_date.'-90 days');
		$ninetyplus_date = date('Y-m-d',$ninetyplus_date);

		$sevenplus_date = strtotime($end_date.'+7 days');
		$sevenplus_date = date('Y-m-d',$sevenplus_date);
		print_r($thirty_date);
		echo "\n";
		print_r($sixty_date);
		echo "\n";
		print_r($ninety_date);
		echo "\n";
		print_r($ninetyplus_date);
		echo "\n";
		print_r($sevenplus_date);
		// exit(0);
		try {
			// testing 
			$loans_due= $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+1 days')))->endkey(date('Y-m-d',strtotime($end_date.'+7 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			$days0to7=0;
			foreach ($loans_due->rows as $row) {
				$days0to7 += $row->value->balance_total;
			}
			// print_r('all after '.$days0to7);
			// print_r($loans_due);
			// exit(0);

			$loan_schedule_year = $this->couchdb->startkey($year_date)->endkey($end_date)->getview('loans', 'get_loan_schedule_by_isodate');
			// print_r(count($loan_schedule_year	->rows));
			// exit(0);
			$loan_schedule = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('loans', 'get_loan_schedule_by_isodate');
			// echo "<br>";
			// print_r(count($loan_schedule->rows));
			// echo "<br>";
			$auction_payments = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('clients', 'get_all_auction_payments_by_isodate');

			// $loans_by_duedate = $this->couchdb->startkey('2015-01-01')->endkey('2015-03-31')->getview('clients', 'get_all_auction_payments_by_isodate');
			$loans_due_after_period = $this->couchdb->startkey(array($end_date,$after_period_date))->endkey(array($end_date,$twelve_months_date))->getview('rbz', 'get_all_loans_by_duedate');
			
			
			//this will replace the before period 30 - 90+
			$loans_due_before_period30 = $this->couchdb->startkey(array($end_date,date('Y-m-d',strtotime($end_date.'-29 days'))))->endkey(array($end_date,$end_date))->getview('rbz', 'get_all_loans_period_open');
			// print_r(count($loans_due_before_period30->rows));
			
			// echo "<br>";
			$loans_due_before_period60 = $this->couchdb->startkey(array($end_date,date('Y-m-d',strtotime($end_date.'-59 days'))))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-30 days'))))->getview('rbz', 'get_all_loans_period_open');
			// print_r(count($loans_due_before_period60->rows));
			// echo "<br>";
			$loans_due_before_period90 = $this->couchdb->startkey(array($end_date,date('Y-m-d',strtotime($end_date.'-89 days'))))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-60 days'))))->getview('rbz', 'get_all_loans_period_open');
			// print_r(count($loans_due_before_period90->rows));
			// echo "<br>";
			$loans_due_before_period90plus = $this->couchdb->startkey(array($end_date,date('Y-m-d',strtotime($end_date.'-500 days'))))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-90 days'))))->getview('rbz', 'get_all_loans_period_open');
			// exit(0);
			// print_r(count($loans_due_after_period->rows));
			$loans_due_before_period_sched6 = $this->couchdb->startkey(array(date('Y-m-d',strtotime($end_date.'+1 days')),$year_date))->endkey(array(date('Y-m-d',strtotime($end_date.'+1 days')),$end_date))->getview('rbz', 'get_all_loans_period_open');
			print_r("all loans before period ".count($loans_due_before_period_sched6->rows));
			echo "\n";
			$loans_due_after_period0to7 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+1 days')))->endkey(date('Y-m-d',strtotime($end_date.'+7 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			print_r(count($loans_due_after_period0to7->rows));
			echo "\n";
			$loans_due_after_period8to14 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+8 days')))->endkey(date('Y-m-d',strtotime($end_date.'+14 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			$loans_due_after_period15to30 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+15 days')))->endkey(date('Y-m-d',strtotime($end_date.'+30 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			$loans_due_after_period31to60 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+31 days')))->endkey(date('Y-m-d',strtotime($end_date.'+60 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			$loans_due_after_period61to90 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+61 days')))->endkey(date('Y-m-d',strtotime($end_date.'+90 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			print_r("61-90 new ".count($loans_due_after_period61to90->rows));
			echo "\n";
			$loans_due_after_period91to120 = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+91 days')))->endkey(date('Y-m-d',strtotime($end_date.'+400 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			// print_r(" ".count($loans_due_before_period_sched6->rows));
			$loans_due_after_test = $this->couchdb->startkey(date('Y-m-d',strtotime($end_date.'+1 days')))->endkey(date('Y-m-d',strtotime($end_date.'+500 days')))->getview('rbz', 'get_all_loans_by_duedate_after_period');
			$loan_after = 0;
			foreach ($loan_schedule_year->rows as $row) {
				$loan_after += $row->value->balance_total;
			}
			print_r("Loan After = ".$loan_after);
			$top_clients_openloans = $this->couchdb->reduce(TRUE)->group(TRUE)->getview('reports', 'get_open_loancount_by_clientid');

			$loans_issued = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('reports', 'get_loan_issue_date');
			$loans_issued_b4_period = $this->couchdb->startkey('2013-01-01')->endkey(date('Y-m-d',strtotime($start_date.'-1 days')))->getview('reports', 'get_loan_issue_date');
			// print_r($loans_due_before_period_sched6);

		// exit(0);
			// $loans_due_before_period30 = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period30');
			// $loans_due_before_period60 = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period60');
			// $loans_due_before_period90 = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period90');
			// $loans_due_before_period90plus = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period90plus');
			// $loans_due_before_period_sched6 = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period_sched6');
			// $loans_due_after_period0to7 = $this->couchdb->startkey('2016-10-01')->endkey('2016-10-07')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			// print_r(count($loans_due_after_period0to7->rows));
			// echo "\n";
			// exit(0);
			// $loans_due_after_period8to14 = $this->couchdb->startkey('2015-04-08')->endkey('2015-04-14')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			// $loans_due_after_period15to30 = $this->couchdb->startkey('2015-04-15')->endkey('2015-04-30')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			// $loans_due_after_period31to60 = $this->couchdb->startkey('2015-05-01')->endkey('2015-05-30')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			 // $loans_due_after_period61to90 = $this->couchdb->startkey('2016-11-30')->endkey('2016-12-29')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			// print_r("61 - 90 ".count($loans_due_after_period61to90->rows));
			// echo "\n";
			// $loans_due_after_period91to120 = $this->couchdb->startkey('2016-06-29')->endkey('2015-07-29')->getview('reports', 'get_all_loans_by_duedate_after_march_sched6');
			// $top_clients_openloans = $this->couchdb->reduce(TRUE)->group(TRUE)->getview('reports', 'get_open_loancount_by_clientid');

			// $loans_issued = $this->couchdb->startkey('2015-01-01')->endkey('2015-03-31')->getview('reports', 'get_loan_issue_date');
			// $loans_issued_b4_period = $this->couchdb->startkey('2013-01-01')->endkey('2015-01-01')->getview('reports', 'get_loan_issue_date');
			// print_r(count($loans_issued->rows));
			// exit(0);
			// print_r($top_clients_openloans);
			// exit(0);
			// print_r($loans_due_after_period91to120);
			// exit(0);
			// print_r($auction_payments);
			// exit(0);
			$total_interest = 0;
			$total_storage = 0;
			$consumer_interest = 0;
			$commercial_interest = 0;
			$other_interest = 0;
			$undefined_interest = 0;
			$consumer_total_loans13 = 0;
			$commercial_total_loans13 = 0;
			$other_total_loans13 = 0;
			// print_r($loan_schedule);
			// exit(0);

			$consumer_array = array('CROSS BORDER TRADERS', 'VENDORS', 'CONSUMPTION', 'FUNERAL ASSISTANCE');
			$commercial_array = array('DISTRIBUTION', 'MANUFACTURING', 'RETAIL', 'SERVICES', 'HEALTH', 'EDUCATION', 'MINING', 'AGRICULTURE', 'CONSTRUCTION', 'TRANSPORT', 'OTHER-WORKING CAPITAL');

			$farming13 = 0;
			$manufacturing13 = 0;
			$mining13 = 0;
			$services13 = 0;
			$retail13 = 0;
			$transport13 = 0;
			$health13 = 0;
			$education13 = 0;
			$distribution13=0;
			$construction13 = 0;
			$consumption13 = 0;
			$cross_border13 = 0;
			$vendors13 = 0;
			$funeral13 = 0;

			$clients_distribution_sched4 = array();
			$clients_manufacturing_sched4 = array();
			$clients_retail_sched4 = array();
			$clients_consumption_sched4 = array();
			$clients_services_sched4 = array();
			$clients_health_sched4 = array();
			$clients_education_sched4 = array();
			$clients_mining_sched4 = array();
			$clients_agriculture_sched4 = array();
			$clients_cross_border_sched4 = array();
			$clients_vendors_sched4 = array();
			$clients_funeral_sched4 = array();
			$clients_other_sched4 = array();
			$loans_distribution_sched4 = 0;
			$loans_manufacturing_sched4 = 0;
			$loans_retail_sched4 = 0;
			$loans_consumption_sched4 = 0;
			$loans_services_sched4 = 0;
			$loans_health_sched4 = 0;
			$loans_education_sched4 = 0;
			$loans_mining_sched4 = 0;
			$loans_agriculture_sched4 = 0;
			$loans_cross_border_sched4 = 0;
			$loans_vendors_sched4 = 0;
			$loans_funeral_sched4 = 0;
			$loans_other_sched4 = 0;

			$f_distribution_sched4 = 0;
			$f_manufacturing_sched4 = 0;
			$f_retail_sched4 = 0;
			$f_consumption_sched4 = 0;
			$f_services_sched4 = 0;
			$f_health_sched4 = 0;
			$f_education_sched4 = 0;
			$f_mining_sched4 = 0;
			$f_agriculture_sched4 = 0;
			$f_cross_border_sched4 = 0;
			$f_vendors_sched4 = 0;
			$f_funeral_sched4 = 0;
			$f_other_sched4 = 0;
			$fclients_distribution_sched4 = array();
			$fclients_manufacturing_sched4 = array();
			$fclients_retail_sched4 = array();
			$fclients_consumption_sched4 = array();
			$fclients_services_sched4 = array();
			$fclients_health_sched4 = array();
			$fclients_education_sched4 = array();
			$fclients_mining_sched4 = array();
			$fclients_agriculture_sched4 = array();
			$fclients_cross_border_sched4 = array();
			$fclients_vendors_sched4 = array();
			$fclients_funeral_sched4 = array();
			$fclients_other_sched4 = array();
			$floans_distribution_sched4 = 0;
			$floans_manufacturing_sched4 = 0;
			$floans_retail_sched4 = 0;
			$floans_consumption_sched4 = 0;
			$floans_services_sched4 = 0;
			$floans_health_sched4 = 0;
			$floans_education_sched4 = 0;
			$floans_mining_sched4 = 0;
			$floans_agriculture_sched4 = 0;
			$floans_cross_border_sched4 = 0;
			$floans_vendors_sched4 = 0;
			$floans_funeral_sched4 = 0;
			$floans_other_sched4 = 0;

			$current_liabilities = 0;
			$underpayments = 0;
			$overpayments = 0;
			foreach ($loan_schedule_year->rows as $row) {
				if ($row->value->balance_interest ===0) {
					$total_interest +=0;
				}else {
					$total_interest += $row->value->loan_daily_interest;
				
				}
				if ($row->value->balance_storage ===0) {
				}else {
					$total_storage += $row->value->loan_daily_storage;
				
				}
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					if ($row->value->balance_interest ===0) {
					} else {
						$consumer_interest += $row->value->loan_daily_interest;
					
					}
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					if ($row->value->balance_interest ===0) {
					} else {
						$commercial_interest += $row->value->loan_daily_interest;
					
					}
					
				} else {
					if ($row->value->balance_interest ===0) {
					} else {
						$other_interest += $row->value->loan_daily_interest;
					
					}
					
				}


			}
			$testtotal = 0;
			$consumertotal = 0;
			$commercialtotal=0;
			foreach ($loan_schedule->rows as $row) {
				if ($row->key == $end_date) {
					$testtotal += $row->value->balance_total;
				}
				/** This is commented out as added above for the whole year not period
				if ($row->value->balance_storage !=0) {
					$total_storage += $row->value->loan_daily_storage;
				}
				**/
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					if ($row->key == $end_date) {
						$consumertotal += $row->value->balance_total;
					}
					/** This is commented out as added above for the whole year not period
					if ($row->value->balance_interest !=0) {
						$consumer_interest += $row->value->loan_daily_interest;
					}
					**/
					if ($row->key == $end_date) {
						$consumer_total_loans13 += $row->value->balance_total;
						$loans_consumption_sched4++;
						array_push($clients_consumption_sched4, $row->value->client->_id);
					}
					if ($row->key == $end_date) {
						if ($row->value->loan_header->purpose_of_loan == 'CROSS BORDER TRADERS') {
							$cross_border13 += $row->value->balance_total;
							$loans_cross_border_sched4++;
							array_push($clients_cross_border_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_cross_border_sched4 += $row->value->balance_total;
								$floans_cross_border_sched4++;
								array_push($fclients_cross_border_sched4, $row->value->client->_id);
							}

						} else if ($row->value->loan_header->purpose_of_loan == 'VENDORS') {
							$vendors13 += $row->value->balance_total;
							$loans_vendors_sched4++;
							array_push($clients_vendors_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_vendors_sched4 += $row->value->balance_total;
								$floans_vendors_sched4++;
								array_push($fclients_vendors_sched4, $row->value->client->_id);
							}

						} else if ($row->value->loan_header->purpose_of_loan == 'FUNERAL ASSISTANCE') {
							$funeral13 += $row->value->balance_total;
							$loans_funeral_sched4++;
							array_push($clients_funeral_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_funeral_sched4 += $row->value->balance_total;
								$floans_funeral_sched4++;
								array_push($fclients_funeral_sched4, $row->value->client->_id);
							}

						} else if ($row->value->loan_header->purpose_of_loan == 'CONSUMPTION') {
							$consumption13 += $row->value->balance_total;
							$loans_consumption_sched4++;
							array_push($clients_consumption_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_consumption_sched4 += $row->value->balance_total;
								$floans_consumption_sched4++;
								array_push($fclients_consumption_sched4, $row->value->client->_id);
							}

						} else {
							echo '<br>' . $row->value->loan_header->purpose_of_loan;
							print_r($row);
							echo "<br>";
						}
					}
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					if ($row->key == $end_date) {
						$commercialtotal += $row->value->balance_total;
					}
					/** This is commented out as added above for the whole year not period
					if ($row->value->balance_interest !=0) {
						$commercial_interest += $row->value->loan_daily_interest;
					}
					**/
					if ($row->key == $end_date) {
						$commercial_total_loans13 += $row->value->balance_total;
						if ($row->value->loan_header->purpose_of_loan == 'AGRICULTURE') {
							$farming13 += $row->value->balance_total;
							$loans_agriculture_sched4++;
							array_push($clients_agriculture_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_agriculture_sched4 += $row->value->balance_total;
								$floans_agriculture_sched4++;
								array_push($fclients_agriculture_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'MANUFACTURING') {
							$manufacturing13 += $row->value->balance_total;
							$loans_manufacturing_sched4++;
							array_push($clients_manufacturing_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_manufacturing_sched4 += $row->value->balance_total;
								$floans_manufacturing_sched4++;
								array_push($fclients_manufacturing_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'MINING') {
							$mining13 += $row->value->balance_total;
							$loans_mining_sched4++;
							array_push($clients_mining_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_mining_sched4 += $row->value->balance_total;
								$floans_mining_sched4++;
								array_push($fclients_mining_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'SERVICES') {
							$services13 += $row->value->balance_total;
							$loans_services_sched4++;
							array_push($clients_services_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_services_sched4 += $row->value->balance_total;
								$floans_services_sched4++;
								array_push($fclients_services_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'RETAIL') {
							$retail13 += $row->value->balance_total;
							$loans_retail_sched4++;
							array_push($clients_retail_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_retail_sched4 += $row->value->balance_total;
								$floans_retail_sched4++;
								array_push($fclients_retail_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'OTHER-WORKING CAPITAL') {
							$retail13 += $row->value->balance_total;
							$loans_retail_sched4++;
							array_push($clients_retail_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_retail_sched4 += $row->value->balance_total;
								$floans_retail_sched4++;
								array_push($fclients_retail_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'TRANSPORT') {
							$transport13 += $row->value->balance_total;

						} else if ($row->value->loan_header->purpose_of_loan == 'HEALTH') {
							$health13 += $row->value->balance_total;
							$loans_health_sched4++;
							array_push($clients_health_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_health_sched4 += $row->value->balance_total;
								$floans_health_sched4++;
								array_push($fclients_health_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'EDUCATION') {
							$education13 += $row->value->balance_total;
							$loans_education_sched4++;
							array_push($clients_education_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_education_sched4 += $row->value->balance_total;
								$floans_education_sched4++;
								array_push($fclients_education_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'DISTRIBUTION') {
							$distribution13 += $row->value->balance_total;
							$loans_distribution_sched4++;
							array_push($clients_distribution_sched4, $row->value->client->_id);
							if ($row->value->client->sex == 'F') {
								$f_distribution_sched4 += $row->value->balance_total;
								$floans_distribution_sched4++;
								array_push($fclients_distribution_sched4, $row->value->client->_id);
							}
						} else if ($row->value->loan_header->purpose_of_loan == 'CONSTRUCTION') {
							$construction13 += $row->value->balance_total;

						} else {
							echo '<br>' . $row->value->loan_header->purpose_of_loan;
						}

					}
				} else {
					/** This is commented out as added above for the whole year not period
					if ($row->value->balance_interest !=0) {
						$other_interest += $row->value->loan_daily_interest;
					}
					**/
					if ($row->key == $end_date) {
						$other_total_loans13 += $row->value->balance_total;
						$loans_other_sched4++;
						array_push($clients_other_sched4, $row->value->client->_id);
						if ($row->value->client->sex == 'F') {
							$f_other_sched4 += $row->value->balance_total;
							$floans_other_sched4++;
							array_push($fclients_other_sched4, $row->value->client->_id);
						}
					}
				}

			}
			echo "\n Consumer Total \n ".$consumertotal;
			echo "\n Commercial Total \n ".$commercialtotal;
			echo "\n Other Total \n ".$other_total_loans13;
			echo "\n Test Total \n ".$testtotal;

			$written_off_count = 0;
			$written_off_value = 0;
			foreach ($auction_payments->rows as $row) {
				$underpayments += $row->value->underpayment;
				$overpayments += $row->value->overpayment;
				if ($row->value->underpayment > 0) {
					$written_off_count++;
					$written_off_value += $row->value->underpayment;
				}
			}
			$current_liabilities = $overpayments - $underpayments;

			$c_consumer_loans = 0;
			$c_commercial_loans = 0;
			$c_other_loans = 0;
			$sm_consumer_loans = 0;
			$sm_commercial_loans = 0;
			$sm_other_loans = 0;
			$ss_consumer_loans = 0;
			$ss_commercial_loans = 0;
			$ss_other_loans = 0;
			$td_consumer_loans = 0;
			$td_commercial_loans = 0;
			$td_other_loans = 0;
			$tl_consumer_loans = 0;
			$tl_commercial_loans = 0;
			$tl_other_loans = 0;

			foreach ($loans_due_after_period->rows as $row) {
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					$c_consumer_loans += $row->value->balance_total;
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					$c_commercial_loans += $row->value->balance_total;
				} else {
					$c_other_loans += $row->value->balance_total;
				}
			}

			foreach ($loans_due_before_period30->rows as $row) {
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					$sm_consumer_loans += $row->value->balance_total;
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					$sm_commercial_loans += $row->value->balance_total;
				} else {
					$sm_other_loans += $row->value->balance_total;
				}
			}

			foreach ($loans_due_before_period60->rows as $row) {
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					$ss_consumer_loans += $row->value->balance_total;
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					$ss_commercial_loans += $row->value->balance_total;
				} else {
					$ss_other_loans += $row->value->balance_total;
				}
			}

			foreach ($loans_due_before_period90->rows as $row) {
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					$td_consumer_loans += $row->value->balance_total;
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					$td_commercial_loans += $row->value->balance_total;
				} else {
					$td_other_loans += $row->value->balance_total;
				}
			}

			foreach ($loans_due_before_period90plus->rows as $row) {
				if (in_array($row->value->loan_header->purpose_of_loan, $consumer_array)) {
					$tl_consumer_loans += $row->value->balance_total;
				} else if (in_array($row->value->loan_header->purpose_of_loan, $commercial_array)) {
					$tl_commercial_loans += $row->value->balance_total;
				} else {
					$tl_other_loans += $row->value->balance_total;
				}
			}

			$days0to7 = 0;
			foreach ($loans_due_before_period_sched6->rows as $row) {
				$days0to7 += $row->value->balance_total;
			}

			foreach ($loans_due_after_period0to7->rows as $row) {
				$days0to7 += $row->value->balance_total;
			}

			$days8to14 = 0;
			foreach ($loans_due_after_period8to14->rows as $row) {
				$days8to14 += $row->value->balance_total;
			}

			$days15to30 = 0;
			foreach ($loans_due_after_period15to30->rows as $row) {
				$days15to30 += $row->value->balance_total;
			}

			$days31to60 = 0;
			foreach ($loans_due_after_period31to60->rows as $row) {
				$days31to60 += $row->value->balance_total;
			}

			$days61to90 = 0;
			foreach ($loans_due_after_period61to90->rows as $row) {
				$days61to90 += $row->value->balance_total;
			}

			$days91to120 = 0;
			foreach ($loans_due_after_period91to120->rows as $row) {
				$days91to120 += $row->value->balance_total;
			}

			$client_array = array();
			$client_array20 = array();
			foreach ($top_clients_openloans->rows as $row) {
				$client_array[$row->key] = $row->value;

			}
			arsort($client_array);
			$x = 0;
			foreach ($client_array as $key => $value) {
				array_push($client_array20, $key);

				// echo "Client " . $key . " " . $value;
				// echo "<br>";
				$x++;
				if ($x == 21) {break;}
			}
			$client_array20_details = array();
			foreach ($client_array20 as $client) {
				print_r($client);
				echo "\n";
				$client_loan = $this->couchdb->key($client)->getview('reports', 'get_open_loans_by_clientid_sched7');
				$client_name = '';
				$client_total = 0;
				$client_balance = 0;
				$client_maturity = '';
				$client_security = '';
				$client_security_value = 0;
				foreach ($client_loan->rows as $row) {
					$client_name = $row->value->client->first_name . " " . $row->value->client->last_name;
					$client_total += $row->value->loan_summary->capital;
					$client_balance += $row->value->balance_total;
					$client_maturity = $row->value->loan_summary->payment_due_date;
					foreach ($row->value->pledge_details as $pledge) {
						$client_security = $pledge->type_of_pledge;
						$client_security_value += $pledge->value;
					}
				}
				$client_array20_details['name'][] = $client_name;
				$client_array20_details['total'][] = $client_total;
				$client_array20_details['balance'][] = $client_balance;
				$client_array20_details['maturity'][] = $client_maturity;
				$client_array20_details['security'][] = $client_security;
				$client_array20_details['security_value'][] = $client_security_value;
			}
			// print_r($client_array20_details);
			// exit(0);

			$loans_issued_count = 0;
			$loans_issued_value = 0;
			$loans_issued_count = count($loans_issued->rows);
			$new_client = 0;
			$existing_client = 0;
			$over5k = 0;
			foreach ($loans_issued->rows as $row) {
				if ($row->value->loan_summary->capital > 5000) {
					$over5k++;
				}
				$loans_issued_value += $row->value->loan_summary->capital;
				$client_id = $row->value->client->_id;
				$exists = false;
				foreach ($loans_issued_b4_period->rows as $issued_row) {
					if ($issued_row->value->client->_id == $client_id) {
						$existing_client++;
						$exists = true;
						break;
					}
				}
				if (!$exists) {
					$new_client++;
				}
			}

			$active_start = $this->couchdb->getview('reports', 'get_open_loans_by_clientid_sched8_begin_period');
			$active_end = $this->couchdb->getview('reports', 'get_open_loans_by_clientid_sched7');
			$active_start_array = array();
			$active_end_array = array();
			foreach ($active_start->rows as $row) {
				array_push($active_start_array, $row->key);
			}
			foreach ($active_end->rows as $row) {
				array_push($active_end_array, $row->key);
			}

			$female_client_perc = 0;
			$all_clients = $this->couchdb->getview('loans', 'get_all_clients_by_id');
			$all_clients_count = count($all_clients->rows);
			$female = 0;
			foreach ($all_clients->rows as $row) {
				if ($row->value->sex == 'F') {
					$female++;
				}
			}
			$female_client_perc = round(($female / $all_clients_count), 2);

			// $loans_by_duedate = $this->couchdb->getview('reports', 'get_all_loans_by_duedate');
			echo("gonna try first one");
			$loans_by_duedate = $this->couchdb->descending(false)->startkey(array($end_date,$year_date))->endkey(array($end_date,$end_date))->getview('rbz', 'get_all_loans_by_duedate');
			echo("loans_by_duedate");
			echo("<br>");
			// $total_loans_due = count($loans_by_duedate->rows);
			$total_loans_due = 0;
			$all_open_loans_count = $total_loans_due;
			$all_open_loans_count += count($loans_due_after_period->rows);

			$total_loans_due_value = 0;
			foreach ($loans_by_duedate->rows as $row) {
				$total_loans_due_value += $row->value->balance_total;
				if ($row->value->balance_capital> 0) {
					$total_loans_due ++;

				}
			}
			// $loans_by_duedate = $this->couchdb->startkey('2013-01-01')->endkey(date('Y-m-d',strtotime($end_date.'-30 days')))->getview('reports', 'get_all_loans_by_duedate');
			echo("gonna try less 30");
			$loans_by_duedate = $this->couchdb->descending(false)->startkey(array($end_date,$year_date))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-30 days'))))->getview('rbz', 'get_all_loans_by_duedate');
			// $total_loans_due30 = count($loans_by_duedate->rows);
			echo("loans_by_duedate less 30");
			echo("<br>");
			$total_loans_due30=0;
			$total_loans_due30_value = 0;
			foreach ($loans_by_duedate->rows as $row) {
				$total_loans_due30_value += $row->value->balance_total;
				if ($row->value->balance_capital> 0) {
					$total_loans_due30 ++;

				}
			}
			// $loans_by_duedate = $this->couchdb->startkey('2013-01-01')->endkey(date('Y-m-d',strtotime($end_date.'-60 days')))->getview('reports', 'get_all_loans_by_duedate');
			$loans_by_duedate = $this->couchdb->descending(false)->startkey(array($end_date,$year_date))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-60 days'))))->getview('rbz', 'get_all_loans_by_duedate');
			// $total_loans_due60 = count($loans_by_duedate->rows);
			echo("loans_by_duedate less 60");
			echo("<br>");
			$total_loans_due60 =0;
			$total_loans_due60_value = 0;
			foreach ($loans_by_duedate->rows as $row) {
				$total_loans_due60_value += $row->value->balance_total;
				if ($row->value->balance_capital> 0) {
					$total_loans_due60 ++;

				}
			}
			// $loans_by_duedate = $this->couchdb->startkey('2013-01-01')->endkey(date('Y-m-d',strtotime($end_date.'-90 days')))->getview('reports', 'get_all_loans_by_duedate');
			$loans_by_duedate = $this->couchdb->descending(true)->startkey(array($end_date,$year_date))->endkey(array($end_date,date('Y-m-d',strtotime($end_date.'-90 days'))))->getview('rbz', 'get_all_loans_by_duedate');
			// $total_loans_due90 = count($loans_by_duedate->rows);
			echo("loans_by_duedate less 90");
			echo("<br>");
			$total_loans_due90 = 0;
			$total_loans_due90_value = 0;
			foreach ($loans_by_duedate->rows as $row) {
				$total_loans_due90_value += $row->value->balance_total;
				if ($row->value->balance_capital> 0) {
					$total_loans_due90 ++;

				}
			}

			$all_payments = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('reports', 'get_all_payments_by_isodate');
			$all_payments_value = 0;
			foreach ($all_payments->rows as $row) {
				$all_payments_value += $row->value->payment_value;
			}

			$all_rolled_loans = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('reports', 'get_all_rolled_loans_by_isodate');
			$all_rolled_value = 0;
			$all_rolled_count = count($all_rolled_loans->rows);
			foreach ($all_rolled_loans->rows as $row) {
				$all_rolled_value += $row->value->rolled_capital;
			}

			$all_loans_before_period = $this->couchdb->getview('reports', 'get_all_loans_by_duedate_before_period_sched8');
			$all_due_during_period = $this->couchdb->startkey($start_date)->endkey($end_date)->getview('reports', 'get_all_loans_by_duedate_loan_summary');
			$total_amount_due_in_period =0;
			foreach ($all_loans_before_period->rows as $row) {
				$total_amount_due_in_period += $row->value->balance_total;
			}
			foreach ($all_due_during_period->rows as $row){
				$total_amount_due_in_period += $row->value->total_repayment_due;
			}

			print_r('Consumer ' . round($consumer_interest, 2));
			echo ('<br>');
			print_r('Commercial ' . round($commercial_interest, 2));
			echo ('<br>');
			print_r('Other ' . round($other_interest, 2));
			echo ('<br>');
			print_r('Undefined ' . round($undefined_interest, 2));
			echo ('<br>');
			print_r('Total ' . round($total_interest, 2));
			echo ('<br>');
			print_r($total_storage / 1.15);
			echo ('<br>');
			print_r('Consumer 13 ' . round($consumer_total_loans13, 2));
			echo ('<br>');
			print_r('Commercial 13 ' . round($commercial_total_loans13, 2));
			echo ('<br>');
			print_r('Other 13 ' . round($other_total_loans13, 2));
			echo ('<br>');
			print_r('Farming13 ' . round($farming13, 2));
			echo ('<br>');
			print_r('manufacturing13 ' . round($manufacturing13, 2));
			echo ('<br>');
			print_r('mining13 ' . round($mining13, 2));
			echo ('<br>');
			print_r('services13 ' . round($services13, 2));
			echo ('<br>');
			print_r('retail13 ' . round($retail13, 2));
			echo ('<br>');
			print_r('transport13 ' . round($transport13, 2));
			echo ('<br>');
			print_r('health13 ' . round($health13, 2));
			echo ('<br>');
			print_r('education13 ' . round($education13, 2));
			echo ('<br>');
			print_r('construction13 ' . round($construction13, 2));
			echo ('<br>');
			print_r('overpayments ' . round($overpayments, 2));
			echo ('<br>');
			print_r('underpayments ' . round($underpayments, 2));
			echo ('<br>');
			print_r('current_liabilities ' . round($current_liabilities, 2));
			echo ('<br>');
			print_r('$c_consumer_loans ' . round($c_consumer_loans, 2));
			echo ('<br>');
			print_r('$c_commercial_loans ' . round($c_commercial_loans, 2));
			echo ('<br>');
			print_r('$c_other_loans ' . round($c_other_loans, 2));
			echo ('<br>');
			print_r('$sm_consumer_loans ' . round($sm_consumer_loans, 2));
			echo ('<br>');
			print_r('$sm_commercial_loans ' . round($sm_commercial_loans, 2));
			echo ('<br>');
			print_r('$sm_other_loans ' . round($sm_other_loans, 2));
			echo ('<br>');
			print_r('$ss_consumer_loans ' . round($ss_consumer_loans, 2));
			echo ('<br>');
			print_r('$ss_commercial_loans ' . round($ss_commercial_loans, 2));
			echo ('<br>');
			print_r('$ss_other_loans ' . round($ss_other_loans, 2));
			echo ('<br>');
			print_r('$td_consumer_loans ' . round($td_consumer_loans, 2));
			echo ('<br>');
			print_r('$td_commercial_loans ' . round($td_commercial_loans, 2));
			echo ('<br>');
			print_r('$td_other_loans ' . round($td_other_loans, 2));
			echo ('<br>');
			print_r('$tl_consumer_loans ' . round($tl_consumer_loans, 2));
			echo ('<br>');
			print_r('$tl_commercial_loans ' . round($tl_commercial_loans, 2));
			echo ('<br>');
			print_r('$tl_other_loans ' . round($tl_other_loans, 2));
			$ttl_sched3 = $c_consumer_loans + $c_commercial_loans + $c_other_loans + $sm_consumer_loans + $sm_commercial_loans + $sm_other_loans + $ss_consumer_loans + $ss_commercial_loans +
			$ss_other_loans + $td_consumer_loans + $td_commercial_loans + $td_other_loans + $tl_consumer_loans + $tl_commercial_loans + $tl_other_loans;
			echo ('<br>');
			print_r('$ttl_sched3 ' . round($ttl_sched3, 2));
			echo ('<br>');
			print_r('$loans_manufacturing_sched4 ' . round($loans_manufacturing_sched4, 2));
			echo ('<br>');
			print_r('$clients_manufacturing_sched4 ' . count(array_unique($clients_manufacturing_sched4)));
			echo ('<br>');
			print_r('$loans_retail_sched4 ' . round($loans_retail_sched4, 2));
			echo ('<br>');
			print_r('$clients_retail_sched4 ' . count(array_unique($clients_retail_sched4)));
			echo ('<br>');
			print_r('$loans_consumption_sched4 ' . round($loans_consumption_sched4, 2));
			echo ('<br>');
			print_r('$clients_consumption_sched4 ' . count(array_unique($clients_consumption_sched4)));
			echo ('<br>');
			print_r('$loans_services_sched4 ' . round($loans_services_sched4, 2));
			echo ('<br>');
			print_r('$clients_services_sched4 ' . count(array_unique($clients_services_sched4)));
			echo ('<br>');
			print_r('$loans_health_sched4 ' . round($loans_health_sched4, 2));
			echo ('<br>');
			print_r('$clients_health_sched4 ' . count(array_unique($clients_health_sched4)));
			echo ('<br>');
			print_r('$loans_education_sched4 ' . round($loans_education_sched4, 2));
			echo ('<br>');
			print_r('$clients_education_sched4 ' . count(array_unique($clients_education_sched4)));
			echo ('<br>');
			print_r('$loans_mining_sched4 ' . round($loans_mining_sched4, 2));
			echo ('<br>');
			print_r('$clients_mining_sched4 ' . count(array_unique($clients_mining_sched4)));
			echo ('<br>');
			print_r('$loans_agriculture_sched4 ' . round($loans_agriculture_sched4, 2));
			echo ('<br>');
			print_r('$clients_agriculture_sched4 ' . count(array_unique($clients_agriculture_sched4)));
			echo ('<br>');
			print_r('$loans_cross_border_sched4 ' . round($loans_cross_border_sched4, 2));
			echo ('<br>');
			print_r('$clients_cross_border_sched4 ' . count(array_unique($clients_cross_border_sched4)));
			echo ('<br>');
			print_r('$loans_vendors_sched4 ' . round($loans_vendors_sched4, 2));
			echo ('<br>');
			print_r('$clients_vendors_sched4 ' . count(array_unique($clients_vendors_sched4)));
			echo ('<br>');
			print_r('$loans_funeral_sched4 ' . round($loans_funeral_sched4, 2));
			echo ('<br>');
			print_r('$clients_funeral_sched4 ' . count(array_unique($clients_funeral_sched4)));
			echo ('<br>');
			print_r('$loans_other_sched4 ' . round($loans_other_sched4, 2));
			echo ('<br>');
			print_r('$clients_other_sched4 ' . count(array_unique($clients_other_sched4)));
			echo ('<br>');
			print_r('$f_manufacturing_sched4 ' . round($f_manufacturing_sched4, 2));
			echo ('<br>');
			print_r('$floans_manufacturing_sched4 ' . round($floans_manufacturing_sched4, 2));
			echo ('<br>');
			print_r('$fclients_manufacturing_sched4 ' . count(array_unique($fclients_manufacturing_sched4)));
			echo ('<br>');
			print_r('$f_retail_sched4 ' . round($f_retail_sched4, 2));
			echo ('<br>');
			print_r('$floans_retail_sched4 ' . round($floans_retail_sched4, 2));
			echo ('<br>');
			print_r('$fclients_retail_sched4 ' . count(array_unique($fclients_retail_sched4)));
			echo ('<br>');
			print_r('$f_consumption_sched4 ' . round($f_consumption_sched4, 2));
			echo ('<br>');
			print_r('$floans_consumption_sched4 ' . round($floans_consumption_sched4, 2));
			echo ('<br>');
			print_r('$fclients_consumption_sched4 ' . count(array_unique($fclients_consumption_sched4)));
			echo ('<br>');
			print_r('$f_services_sched4 ' . round($f_services_sched4, 2));
			echo ('<br>');
			print_r('$floans_services_sched4 ' . round($floans_services_sched4, 2));
			echo ('<br>');
			print_r('$fclients_services_sched4 ' . count(array_unique($fclients_services_sched4)));
			echo ('<br>');
			print_r('$f_health_sched4 ' . round($f_health_sched4, 2));
			echo ('<br>');
			print_r('$floans_health_sched4 ' . round($floans_health_sched4, 2));
			echo ('<br>');
			print_r('$fclients_health_sched4 ' . count(array_unique($fclients_health_sched4)));
			echo ('<br>');
			print_r('$f_education_sched4 ' . round($f_education_sched4, 2));
			echo ('<br>');
			print_r('$floans_education_sched4 ' . round($floans_education_sched4, 2));
			echo ('<br>');
			print_r('$fclients_education_sched4 ' . count(array_unique($fclients_education_sched4)));
			echo ('<br>');
			print_r('$f_mining_sched4 ' . round($f_mining_sched4, 2));
			echo ('<br>');
			print_r('$floans_mining_sched4 ' . round($floans_mining_sched4, 2));
			echo ('<br>');
			print_r('$fclients_mining_sched4 ' . count(array_unique($fclients_mining_sched4)));
			echo ('<br>');
			print_r('$f_agriculture_sched4 ' . round($f_agriculture_sched4, 2));
			echo ('<br>');
			print_r('$floans_agriculture_sched4 ' . round($floans_agriculture_sched4, 2));
			echo ('<br>');
			print_r('$fclients_agriculture_sched4 ' . count(array_unique($fclients_agriculture_sched4)));
			echo ('<br>');
			print_r('$f_cross_border_sched4 ' . round($f_cross_border_sched4, 2));
			echo ('<br>');
			print_r('$floans_cross_border_sched4 ' . round($floans_cross_border_sched4, 2));
			echo ('<br>');
			print_r('$fclients_cross_border_sched4 ' . count(array_unique($fclients_cross_border_sched4)));
			echo ('<br>');
			print_r('$f_vendors_sched4 ' . round($f_vendors_sched4, 2));
			echo ('<br>');
			print_r('$floans_vendors_sched4 ' . round($floans_vendors_sched4, 2));
			echo ('<br>');
			print_r('$fclients_vendors_sched4 ' . count(array_unique($fclients_vendors_sched4)));
			echo ('<br>');
			print_r('$f_funeral_sched4 ' . round($f_funeral_sched4, 2));
			echo ('<br>');
			print_r('$floans_funeral_sched4 ' . round($floans_funeral_sched4, 2));
			echo ('<br>');
			print_r('$fclients_funeral_sched4 ' . count(array_unique($fclients_funeral_sched4)));
			echo ('<br>');
			print_r('$f_other_sched4 ' . round($f_other_sched4, 2));
			echo ('<br>');
			print_r('$floans_other_sched4 ' . round($floans_other_sched4, 2));
			echo ('<br>');
			print_r('$fclients_other_sched4 ' . count(array_unique($fclients_other_sched4)));
			echo ('<br>');
			print_r('$days0to7 ' . round($days0to7, 2));
			echo ('<br>');
			print_r('$days8to14 ' . round($days8to14, 2));
			echo ('<br>');
			print_r('$days15to30 ' . round($days15to30, 2));
			echo ('<br>');
			print_r('$days31to60 ' . round($days31to60, 2));
			echo ('<br>');
			print_r('$days61to90 ' . round($days61to90, 2));
			echo ('<br>');
			print_r('$days91to120 ' . round($days91to120, 2));
			echo ('<br>');
			print_r('$loans_issued_count ' . $loans_issued_count);
			echo ('<br>');
			print_r('$loans_issued_value ' . round($loans_issued_value, 2));
			echo ('<br>');
			print_r('$new_client ' . $new_client);
			echo ('<br>');
			print_r('existing_client ' . $existing_client);
			echo '<br>';
			echo '$active_start ' . count($active_start_array);
			echo '<br>';
			echo '$active_start unique ' . count(array_unique($active_start_array));
			echo '<br>';
			echo '$active_end ' . count($active_end_array);
			echo '<br>';
			echo '$active_end unique ' . count(array_unique($active_end_array));
			echo '<br>';
			echo '$all_clients_count ' . $all_clients_count;
			echo '<br>';
			echo '$over5k ' . $over5k;
			echo '<br>';
			echo '$female_client_perc ' . $female_client_perc;
			echo '<br>';
			echo '$total_loans_due ' . $total_loans_due;
			echo '<br>';
			echo '$total_loans_due_value ' . round($total_loans_due_value, 2);
			echo '<br>';
			echo '$total_loans_due30 ' . $total_loans_due30;
			echo '<br>';
			echo '$total_loans_due30_value ' . round($total_loans_due30_value, 2);
			echo '<br>';
			echo '$total_loans_due60 ' . $total_loans_due60;
			echo '<br>';
			echo '$total_loans_due60_value ' . round($total_loans_due60_value, 2);
			echo '<br>';
			echo '$total_loans_due90 ' . $total_loans_due90;
			echo '<br>';
			echo '$total_loans_due90_value ' . round($total_loans_due90_value, 2);
			echo '<br>';
			echo '$all_payments_value ' . round($all_payments_value, 2);
			echo '<br>';
			echo '$all_rolled_count ' . $all_rolled_count;
			echo '<br>';
			echo '$all_rolled_value ' . round($all_rolled_value, 2);
			echo '<br>';
			echo '$written_off_count ' . $written_off_count;
			echo '<br>';
			echo '$written_off_value111111 ' . round($written_off_value, 2);
			echo '<br>';
			echo '$total_amount_due_in_period ' . round($total_amount_due_in_period, 2);

			$objReader = PHPExcel_IOFactory::createReader('Excel2007');
			$folder_path = dirname(dirname(dirname(__FILE__)));


			$objPHPExcel = $objReader->load($folder_path.'/rbz.xlsx');
			$objPHPExcel->setActiveSheetIndex(1)
			            ->setCellValue('C12', round($consumer_interest, 2))
			            ->setCellValue('C13', round($commercial_interest, 2))
			            ->setCellValue('C14', round($other_interest, 2))
			            ->setCellValue('C27', round(($total_storage / 1.15), 2))
			            ->setCellValue('D27', round(($total_storage), 2));
			$objPHPExcel->setActiveSheetIndex(2)
			            ->setCellValue('C20', round($consumer_total_loans13, 2))
			            ->setCellValue('C22', round($farming13, 2))
			            ->setCellValue('C23', round($manufacturing13, 2))
			            ->setCellValue('C24', round($mining13, 2))
			            ->setCellValue('C25', round($services13, 2))
			            ->setCellValue('C26', round($retail13, 2))
			            ->setCellValue('C27', round($transport13, 2))
			            ->setCellValue('C28', round($health13, 2))
			            ->setCellValue('C29', round($education13, 2))
			            ->setCellValue('C30', round($construction13, 2))
			            ->setCellValue('C31', round($distribution13, 2))
			            ->setCellValue('C32', round($other_total_loans13, 2));
			$objPHPExcel->setActiveSheetIndex(5)
			            ->setCellValue('C13', round($c_consumer_loans, 2))
			            ->setCellValue('C14', round($c_commercial_loans, 2))
			            ->setCellValue('C15', round($c_other_loans, 2))
			            ->setCellValue('C17', round($sm_consumer_loans, 2))
			            ->setCellValue('C18', round($sm_commercial_loans, 2))
			            ->setCellValue('C19', round($sm_other_loans, 2))
			            ->setCellValue('C21', round($ss_consumer_loans, 2))
			            ->setCellValue('C22', round($ss_commercial_loans, 2))
			            ->setCellValue('C23', round($ss_other_loans, 2))
			            ->setCellValue('C25', round($td_consumer_loans, 2))
			            ->setCellValue('C26', round($td_commercial_loans, 2))
			            ->setCellValue('C27', round($td_other_loans, 2))
			            ->setCellValue('C29', round($tl_consumer_loans, 2))
			            ->setCellValue('C30', round($tl_commercial_loans, 2))
			            ->setCellValue('C31', round($tl_other_loans, 2));
			$objPHPExcel->setActiveSheetIndex(6)
						->setCellValue('B11', count(array_unique($clients_distribution_sched4)))
			            ->setCellValue('C11', round($loans_distribution_sched4, 2))
			            ->setCellValue('D11', round($distribution13, 2))
			            ->setCellValue('B12', count(array_unique($clients_manufacturing_sched4)))
			            ->setCellValue('C12', round($loans_manufacturing_sched4, 2))
			            ->setCellValue('D12', round($manufacturing13, 2))
			            ->setCellValue('B13', count(array_unique($clients_retail_sched4)))
			            ->setCellValue('C13', round($loans_retail_sched4, 2))
			            ->setCellValue('D13', round($retail13, 2))
			            ->setCellValue('B14', count(array_unique($clients_consumption_sched4)))
			            ->setCellValue('C14', round($loans_consumption_sched4, 2))
			            ->setCellValue('D14', round($consumption13, 2))
			            ->setCellValue('B15', count(array_unique($clients_services_sched4)))
			            ->setCellValue('C15', round($loans_services_sched4, 2))
			            ->setCellValue('D15', round($services13, 2))
			            ->setCellValue('B16', count(array_unique($clients_health_sched4)))
			            ->setCellValue('C16', round($loans_health_sched4, 2))
			            ->setCellValue('D16', round($health13, 2))
			            ->setCellValue('B17', count(array_unique($clients_education_sched4)))
			            ->setCellValue('C17', round($loans_education_sched4, 2))
			            ->setCellValue('D17', round($education13, 2))
			            ->setCellValue('B18', count(array_unique($clients_mining_sched4)))
			            ->setCellValue('C18', round($loans_mining_sched4, 2))
			            ->setCellValue('D18', round($mining13, 2))
			            ->setCellValue('B19', count(array_unique($clients_agriculture_sched4)))
			            ->setCellValue('C19', round($loans_agriculture_sched4, 2))
			            ->setCellValue('D19', round($farming13, 2))
			            ->setCellValue('B20', count(array_unique($clients_cross_border_sched4)))
			            ->setCellValue('C20', round($loans_cross_border_sched4, 2))
			            ->setCellValue('D20', round($cross_border13, 2))
			            ->setCellValue('B21', count(array_unique($clients_vendors_sched4)))
			            ->setCellValue('C21', round($loans_vendors_sched4, 2))
			            ->setCellValue('D21', round($vendors13, 2))
			            ->setCellValue('B22', count(array_unique($clients_funeral_sched4)))
			            ->setCellValue('C22', round($loans_funeral_sched4, 2))
			            ->setCellValue('D22', round($funeral13, 2))
			            ->setCellValue('B23', count(array_unique($clients_other_sched4)))
			            ->setCellValue('C23', round($loans_other_sched4, 2))
			            ->setCellValue('D23', round($other_total_loans13, 2));
			$objPHPExcel->setActiveSheetIndex(7)
						->setCellValue('B11', count(array_unique($fclients_distribution_sched4)))
			            ->setCellValue('C11', round($floans_distribution_sched4, 2))
			            ->setCellValue('D11', round($f_distribution_sched4, 2))
			            ->setCellValue('B12', count(array_unique($fclients_manufacturing_sched4)))
			            ->setCellValue('C12', round($floans_manufacturing_sched4, 2))
			            ->setCellValue('D12', round($f_manufacturing_sched4, 2))
			            ->setCellValue('B13', count(array_unique($fclients_retail_sched4)))
			            ->setCellValue('C13', round($floans_retail_sched4, 2))
			            ->setCellValue('D13', round($f_retail_sched4, 2))
			            ->setCellValue('B14', count(array_unique($fclients_consumption_sched4)))
			            ->setCellValue('C14', round($floans_consumption_sched4, 2))
			            ->setCellValue('D14', round($f_consumption_sched4, 2))
			            ->setCellValue('B15', count(array_unique($fclients_services_sched4)))
			            ->setCellValue('C15', round($floans_services_sched4, 2))
			            ->setCellValue('D16', round($f_services_sched4, 2))
			            ->setCellValue('B17', count(array_unique($fclients_health_sched4)))
			            ->setCellValue('C17', round($floans_health_sched4, 2))
			            ->setCellValue('D17', round($f_health_sched4, 2))
			            ->setCellValue('B18', count(array_unique($fclients_education_sched4)))
			            ->setCellValue('C18', round($floans_education_sched4, 2))
			            ->setCellValue('D18', round($f_education_sched4, 2))
			            ->setCellValue('B19', count(array_unique($fclients_mining_sched4)))
			            ->setCellValue('C19', round($floans_mining_sched4, 2))
			            ->setCellValue('D19', round($f_mining_sched4, 2))
			            ->setCellValue('B20', count(array_unique($fclients_agriculture_sched4)))
			            ->setCellValue('C20', round($floans_agriculture_sched4, 2))
			            ->setCellValue('D20', round($f_agriculture_sched4, 2))
			            ->setCellValue('B21', count(array_unique($fclients_cross_border_sched4)))
			            ->setCellValue('C21', round($floans_cross_border_sched4, 2))
			            ->setCellValue('D21', round($f_cross_border_sched4, 2))
			            ->setCellValue('B22', count(array_unique($fclients_funeral_sched4)))
			            ->setCellValue('C22', round($floans_funeral_sched4, 2))
			            ->setCellValue('D22', round($f_funeral_sched4, 2))
			            ->setCellValue('B23', count(array_unique($fclients_other_sched4)))
			            ->setCellValue('C23', round($floans_other_sched4, 2))
			            ->setCellValue('D23', round($f_other_sched4, 2));
			$objPHPExcel->setActiveSheetIndex(8)
			            ->setCellValue('B12', round($days0to7, 2))
			            ->setCellValue('B13', round($days8to14, 2))
			            ->setCellValue('B14', round($days15to30, 2))
			            ->setCellValue('B15', round($days31to60, 2))
			            ->setCellValue('B16', round($days61to90, 2))
			            ->setCellValue('B17', round($days91to120, 2));
			print_r('$client_array20_details '.count($client_array20_details));
			for ($x = 0; $x < 21; $x++) {

				$row = 11 + $x;
				$objPHPExcel->setActiveSheetIndex(9)->setCellValue('A' . $row, $client_array20_details['name'][$x])
				            ->setCellValue('A' . $row, $client_array20_details['name'][$x])
				            ->setCellValue('B' . $row, $client_array20_details['total'][$x])
				            ->setCellValue('C' . $row, $client_array20_details['balance'][$x])
				            ->setCellValue('D' . $row, $client_array20_details['maturity'][$x])
				            ->setCellValue('E' . $row, $client_array20_details['security'][$x])
				            ->setCellValue('F' . $row, $client_array20_details['security_value'][$x]);
			}
			$objPHPExcel->setActiveSheetIndex(10)
			            ->setCellValue('C14', $loans_issued_count)
			            ->setCellValue('C15', round($loans_issued_value, 2))
			            ->setCellValue('C16', $new_client)
			            ->setCellValue('C17', $existing_client)
			            ->setCellValue('C18', count(array_unique($active_start_array)))
			            ->setCellValue('C19', count(array_unique($active_end_array)))
			            ->setCellValue('C21', $all_open_loans_count)
			            ->setCellValue('C22', 3)
			            ->setCellValue('C23', $all_clients_count)
			            ->setCellValue('C32', $over5k)
			            ->setCellValue('C33', $female_client_perc)
			            ->setCellValue('C35', $total_loans_due)
			            ->setCellValue('C36', round($total_loans_due_value, 2))
			            ->setCellValue('C37', $total_loans_due30)
			            ->setCellValue('C38', round($total_loans_due30_value, 2))
			            ->setCellValue('C39', $total_loans_due60)
			            ->setCellValue('C40', round($total_loans_due60_value, 2))
			            ->setCellValue('C41', $total_loans_due90)
			            ->setCellValue('C42', round($total_loans_due90_value, 2))
			            ->setCellValue('C43', round($all_payments_value, 2))
			            ->setCellValue('C44', round($total_amount_due_in_period, 2))
			            ->setCellValue('C46', round($all_rolled_value, 2))
			            ->setCellValue('C47', $all_rolled_count)
			            ->setCellValue('C49', $written_off_count)
			            ->setCellValue('C50', round($written_off_value, 2));

			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
			$objWriter->save('/tmp/rbz.xlsx');
		} catch (Exception $e) {
			print_r($e);
		}
	}
}