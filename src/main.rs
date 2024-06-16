use chrono::Duration;
use etender::Op;
use std::thread;
fn main() {
    let signal_path = etender::get_signal_path();
    etender::clear_old_files(3);
    let args: Vec<String> = std::env::args().collect();
    let mut op = Op::Production;
    for arg in &args {
        if arg.to_lowercase() == "debug" {
            println!("debug mode");
            op = Op::Debug;
        }
    }
    etender::receive_messages(&op, &signal_path);
    match etender::send_tender_documents(&op, &signal_path) {
        Ok(_) => (),
        Err(e) => etender::send_message(
            &Op::Debug,
            format!("tender docs not found. {}", e),
            &signal_path,
        ),
    }
    thread::sleep(Duration::milliseconds(100).to_std().unwrap());
    match etender::find_tenders(&op, &signal_path) {
        Ok(_) => (),
        Err(e) => etender::send_message(
            &Op::Debug,
            format!("unable to scrape for tenders. {}", e),
            &signal_path,
        ),
    };
}
