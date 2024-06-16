use calamine::{open_workbook, Reader, Xls};
use chrono::{Days, Duration, Local, NaiveDate};
use docx_rs::*;
use num_format::{Buffer, Locale};
use regex::{Regex, RegexBuilder};
use reqwest::{
    blocking::{Client, Response},
    Error,
};
use rusty_tesseract::image::{DynamicImage, GenericImageView, Pixel};
use std::{
    collections::{HashMap, HashSet},
    error,
    fs::{self, File},
    io::{self, BufRead, BufReader, Write},
    path::Path,
    process::Command,
    str,
};

#[derive(Clone)]
pub enum Op {
    Debug,
    Production,
}

const FOLDER_PREFIX: &str = "/home/jerin/RustProjects/etender/tenders/";
const SIGNAL_DEV_NUMBER: &str = "+919447880206";
const SIGNAL_SENDER_NUMBER: &str = "+919074221997";
const SIGNAL_USER_NUMBER: &str = "+919400507801";
const SIGNAL_CONFIG_PATH: &str = "/home/jerin/.local/share/signal-cli";
const BASE_URL: &str = "https://etenders.kerala.gov.in/nicgep/app";
const EKM_COST_INDEX: f64 = 1.3559;
const APP_USER_AGENT: &str = concat!(env!("CARGO_PKG_NAME"), "/", env!("CARGO_PKG_VERSION"),);

pub fn find_tenders(op: &Op, signal_path: &str) -> Result<(), Box<dyn error::Error>> {
    println!("{}", Local::now());
    let folder = FOLDER_PREFIX.to_string() + Local::now().format("%d-%b-%y").to_string().as_ref();
    let mut folder = Path::new(&folder);
    if !folder.exists() {
        println!("Creating folder.");
        fs::create_dir(folder).expect("unable to create directory");
    }
    let mut modified_folder_path = folder.to_path_buf();
    match op {
        Op::Debug => {
            modified_folder_path = folder.join("debug");
            if !modified_folder_path.exists() {
                fs::create_dir(&modified_folder_path).expect("unable to create directory");
            }
        }
        Op::Production => (),
    };
    folder = &modified_folder_path;
    let txtpath = folder.join("tenders.txt");
    if txtpath.exists() {
        match op {
            Op::Debug => (),
            Op::Production => {
                return Ok(());
            }
        }
    }
    println!("Looking for tenders...");
    let mut message = String::new();
    let mut para = vec![];
    let mut client = Client::builder().cookie_store(true).build().unwrap();
    let _res = get_url_response(&client, BASE_URL);

    let url = BASE_URL.to_string()
        + "?component=%24DirectLink&page=FrontEndTendersByOrganisation&service=direct";
    let res = get_url_response(&client, &url)?;
    let res = res.text()?;
    println!("Parsing  tenders...");
    let re = Regex::new("<td align=\"center\">(.*?)</td>\\s+<td align=\"center\">.*?</td>\\s+<!-- <td align=\"center\"><span.*?/></td> -->\\s+<td align=\"center\"><a id=\".*?\" title=\"View Tender Information\" href=\".*?sp=(.*?)\">(.*?)</a>\\s+.*?\\[.*?\\]\\[(.*?)\\]\\s+</td>\\s+<td align=\"center\">Local Self Government Department.*?Ernakulam(.*?)Panchayath.*?\\|\\|(.*?)</td>")?;
    let mut i = 1;
    para.push(create_docx_para("LSGD Ernakulam", true));
    for cap in re.captures_iter(&res) {
        message += &format!("{}. {}\n", i, &cap[2]);
        let amount = get_tender_amount(&client, &cap[2])?;
        para.push(create_docx_para(
            format!("{}.  {}", i, &cap[3]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Tendering Authority: {}", &cap[6]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Closing Date: {}", &cap[1]).as_str(),
            false,
        ));
        if amount > 0.0 {
            let mut buf = Buffer::default();
            buf.write_formatted(&(amount as i32), &Locale::en_IN);
            para.push(create_docx_para(
                format!("Amount: {}", buf.as_str()).as_str(),
                false,
            ));
        }
        i += 1;
    }
    let mut docxobject = Docx::new();
    for parag in &para {
        docxobject = docxobject.add_paragraph(parag.to_owned());
    }
    let docxpath = folder.join("LSGD.docx");
    let file = File::create(&docxpath).unwrap();
    docxobject.build().pack(file)?;
    para.clear();
    send_attachment(op, &docxpath, signal_path)?;
    let re = Regex::new("<td align=\"center\">(.*?)</td>\\s+<td align=\"center\">.*?</td>\\s+<!-- <td align=\"center\"><span.*?/></td> -->\\s+<td align=\"center\"><a id=\".*?\" title=\"View Tender Information\" href=\".*?sp=(.*?)\">(.*?)</a>\\s+.*?\\[.*?\\]\\[(.*?)\\]\\s+</td>\\s+<td align=\"center\">PWD.*?Buildings.*?Bldgs \\(C\\)(.*?)</td>")?;
    para.push(create_docx_para("PWD", true));
    for cap in re.captures_iter(&res) {
        message += &format!("{}. {}\n", i, &cap[2]);
        let amount = match get_tender_amount(&client, &cap[2]) {
            Ok(amt) => amt,
            Err(_) => {
                client = client.clone();
                get_tender_amount(&client, &cap[2])?
            }
        };
        para.push(create_docx_para(
            format!("{}.  {}", i, &cap[3]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Tendering Authority: SE, Bldgs (C){}", &cap[5]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Closing Date: {}", &cap[1]).as_str(),
            false,
        ));
        if amount > 0.0 {
            let mut buf = Buffer::default();
            buf.write_formatted(&(amount as i32), &Locale::en_IN);
            para.push(create_docx_para(
                format!("Amount: {}", buf.as_str()).as_str(),
                false,
            ));
        }
        i += 1;
    }
    let mut docxobject = Docx::new();
    for parag in &para {
        docxobject = docxobject.add_paragraph(parag.to_owned());
    }
    let docxpath = folder.join("PWD.docx");
    let file = File::create(&docxpath).unwrap();
    docxobject.build().pack(file)?;
    para.clear();
    send_attachment(op, &docxpath, signal_path)?;
    para.push(create_docx_para("MIC Ernakulam", true));
    let re = Regex::new("<td align=\"center\">(.*?)</td>\\s+<td align=\"center\">.*?</td>\\s+<!-- <td align=\"center\"><span.*?/></td> -->\\s+<td align=\"center\"><a id=\".*?\" title=\"View Tender Information\" href=\".*?sp=(.*?)\">(.*?)</a>\\s+.*?\\[.*?\\]\\[(.*?)\\]\\s+</td>\\s+<td align=\"center\">Irrigation.*?SE,MIC,EKLM(.*?)</td>")?;
    for cap in re.captures_iter(&res) {
        message += &format!("{}. {}\n", i, &cap[2]);
        let amount = get_tender_amount(&client, &cap[2])?;
        para.push(create_docx_para(
            format!("{}.  {}", i, &cap[3]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Tendering Authority: SE, MIC, EKLM{}", &cap[5]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Closing Date: {}", &cap[1]).as_str(),
            false,
        ));
        if amount > 0.0 {
            let mut buf = Buffer::default();
            buf.write_formatted(&(amount as i32), &Locale::en_IN);
            para.push(create_docx_para(
                format!("Amount: {}", buf.as_str()).as_str(),
                false,
            ));
        }
        i += 1;
    }
    let mut docxobject = Docx::new();
    for parag in &para {
        docxobject = docxobject.add_paragraph(parag.to_owned());
    }
    let docxpath = folder.join("MIC.docx");
    let file = File::create(&docxpath).unwrap();
    docxobject.build().pack(file)?;
    para.clear();
    send_attachment(op, &docxpath, signal_path)?;
    para.push(create_docx_para("Project Circle Piravom", true));
    let re = Regex::new("<td align=\"center\">(.*?)</td>\\s+<td align=\"center\">.*?</td>\\s+<!-- <td align=\"center\"><span.*?/></td> -->\\s+<td align=\"center\"><a id=\".*?\" title=\"View Tender Information\" href=\".*?sp=(.*?)\">(.*?)</a>\\s+.*?\\[.*?\\]\\[(.*?)\\]\\s+</td>\\s+<td align=\"center\">Irrigation.*?SE,Proj.Cir., Piravom(.*?)</td>")?;
    for cap in re.captures_iter(&res) {
        message += &format!("{}. {}\n", i, &cap[2]);
        let amount = get_tender_amount(&client, &cap[2])?;
        para.push(create_docx_para(
            format!("{}.  {}", i, &cap[3]).as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!(
                "Tendering Authority: SE, Project Circle, Piravom{}",
                &cap[5]
            )
            .as_str(),
            false,
        ));
        para.push(create_docx_para(
            format!("Closing Date: {}", &cap[1]).as_str(),
            false,
        ));
        if amount > 0.0 {
            let mut buf = Buffer::default();
            buf.write_formatted(&(amount as i32), &Locale::en_IN);
            para.push(create_docx_para(
                format!("Amount: {}", buf.as_str()).as_str(),
                false,
            ));
        }
        i += 1;
    }
    let mut docxobject = Docx::new();
    for parag in &para {
        docxobject = docxobject.add_paragraph(parag.to_owned());
    }
    let docxpath = folder.join("PCP.docx");
    let file = File::create(&docxpath).unwrap();
    docxobject.build().pack(file)?;
    send_attachment(op, &docxpath, signal_path)?;
    send_message(
        op,
        format!(
            "Send a number between {} and {} to get the corresponding tender schedule",
            1,
            i - 1
        ),
        signal_path,
    );
    fs::write(txtpath, message)?;
    Ok(())
}

pub fn send_message(op: &Op, message: String, signal_path: &str) {
    let tonumber = match op {
        Op::Debug => SIGNAL_DEV_NUMBER,
        Op::Production => SIGNAL_USER_NUMBER,
    };
    Command::new(&signal_path)
        .arg("--config")
        .arg(SIGNAL_CONFIG_PATH)
        .arg("-a")
        .arg(SIGNAL_SENDER_NUMBER)
        .arg("send")
        .arg(tonumber)
        .arg("-m")
        .arg(&message)
        .output()
        .expect("sending failed");
    println!("Message sent: {}", message);
}

fn send_attachment(op: &Op, path: &Path, signal_path: &str) -> Result<(), Box<dyn error::Error>> {
    let tonumber = match op {
        Op::Debug => SIGNAL_DEV_NUMBER,
        Op::Production => SIGNAL_USER_NUMBER,
    };
    let mut i = 0;
    loop {
        match Command::new(&signal_path)
            .arg("--config")
            .arg(SIGNAL_CONFIG_PATH)
            .arg("-a")
            .arg(SIGNAL_SENDER_NUMBER)
            .arg("send")
            .arg(tonumber)
            .arg("-a")
            .arg(path)
            .output()
        {
            Err(e) => {
                i += 1;
                println!("{}", e);
                std::thread::sleep(Duration::seconds(5).to_std().unwrap());
                if i == 5 {
                    send_message(
                        &Op::Debug,
                        "Sending attachment keeps failing".to_string(),
                        signal_path
                    );
                    return Err("Sending attachment failed!".into());
                }
            }
            Ok(o) => {
                println!("{:?}", o);
                println!("Attachment sent: {}", path.display());
                return Ok(());
            }
        }
    }
}

pub fn receive_messages(op: &Op, signal_path: &str) {
    println!("{}", Local::now());
    println!("Looking for messages...");
    let out = match Command::new(&signal_path)
        .arg("--config")
        .arg(SIGNAL_CONFIG_PATH)
        .arg("-a")
        .arg(SIGNAL_SENDER_NUMBER)
        .arg("receive")
        .output()
    {
        Ok(o) => o,
        Err(e) => {
            println!("{}", e);
            return;
        }
    };
    println!("Received messages.");
    let fromnumber = match op {
        Op::Debug => SIGNAL_DEV_NUMBER,
        Op::Production => SIGNAL_USER_NUMBER,
    };
    let txtpath = FOLDER_PREFIX.to_string() + "remainingqueries.txt";
    let mut f = File::options()
        .append(true)
        .read(true)
        .create(true)
        .open(txtpath)
        .unwrap();
    std::thread::sleep(Duration::seconds(35).to_std().unwrap());
    let outstring = str::from_utf8(&out.stdout).unwrap();
    let messages = outstring.split("Envelope from").collect::<Vec<&str>>();
    for message in messages {
        if message.contains(fromnumber) && message.contains("Body") {
            let messagelines = message.lines();
            for messageline in messagelines {
                let messageline = messageline.trim();
                if let Some(numstr) = messageline.strip_prefix("Body:") {
                    let numstr = numstr.trim();
                    if let Ok(parsedout) = numstr.parse::<usize>() {
                        writeln!(f, "{}", parsedout).unwrap();
                    }
                }
            }
        }
    }
}

fn get_tender_amount(client: &Client, urlsuffix: &str) -> Result<f64, Box<dyn error::Error>> {
    let url = BASE_URL.to_string() + "?component=%24DirectLink_0&page=FrontEndAdvancedSearchResult&service=direct&session=T&sp=" + urlsuffix;
    let res_in = get_url_response(client, &url).unwrap();
    let res_in = res_in.text()?;
    let re_in = Regex::new(".*?Tender Value in.*?\n.*?td_field\">[ ]*?([0-9,]*?)[ ]*?</td>")?;
    let mut amount = -1.0;
    for cap_in in re_in.captures_iter(&res_in) {
        amount = cap_in[1].trim().replace(",", "").parse::<f64>().unwrap();
    }
    Ok(amount)
}

pub fn send_tender_documents(op: &Op, signal_path: &str) -> Result<(), Box<dyn error::Error>> {
    let txtpath = FOLDER_PREFIX.to_string() + "remainingqueries.txt";
    let mut f_qrs = File::options()
        .write(true)
        .read(true)
        .create(true)
        .open(txtpath)
        .unwrap();
    let mut ids = HashSet::new();
    let qrs_reader = BufReader::new(&f_qrs);
    let lines = qrs_reader.lines();
    let linesvec: Vec<io::Result<String>> = lines.collect();
    for line in linesvec {
        let linestring = line.unwrap();
        let linestring = linestring.trim();
        if linestring.is_empty() {
            continue;
        }
        ids.insert(linestring.parse::<usize>().unwrap());
    }
    if ids.is_empty() {
        println!("Empty request.");
        return Ok(());
    }

    let time = Local::now();
    let mut date = time.format("%d-%b-%y").to_string();
    let mut foldername = FOLDER_PREFIX.to_string() + &date;
    let mut folder = Path::new(&foldername);
    let newdate = !folder.exists();
    if newdate {
        date = time
            .checked_sub_days(Days::new(1))
            .unwrap()
            .format("%d-%b-%y")
            .to_string();
        foldername = FOLDER_PREFIX.to_string() + &date;
        folder = Path::new(&foldername);
        if !folder.exists() {
            return Err("folder not found!".into());
        }
    }
    let mut modified_folder_path = folder.to_path_buf();
    match op {
        Op::Debug => {
            modified_folder_path = folder.join("debug");
            if !modified_folder_path.exists() {
                modified_folder_path = folder.to_path_buf();
            }
        }
        Op::Production => (),
    };
    folder = &modified_folder_path;

    let mut sor_map = HashMap::new();
    let mut sor_id_map = HashMap::new();
    let mut sor_opened = false;
    let sor_path = Path::new(FOLDER_PREFIX)
        .join("SOURCEDATA")
        .join("ratesdata.txt");

    let txtpath = folder.join("tenders.txt");
    let f = File::open(txtpath)?;
    let reader = BufReader::new(f);
    let lines = reader.lines();
    let linesvec: Vec<io::Result<String>> = lines.collect();

    let client = Client::builder().cookie_store(true).build().unwrap();
    let _res = get_url_response(&client, BASE_URL);
    let mut succeeded_ids = HashSet::new();

    for idptr in &ids {
        let id = *idptr;
        if id <= linesvec.len() {
            println!("{}", id);
            let linestring = linesvec[id - 1].as_ref().unwrap();
            let v: Vec<&str> = linestring.split(". ").collect();
            assert!(v[0].parse::<usize>().unwrap() == id);
            let amount = get_tender_amount(&client, v[1])?;
            let url = BASE_URL.to_string() + "?component=%24DirectLink_8&page=FrontEndTenderDetails&service=direct&session=T&sp=" + v[1];
            let res = get_url_response(&client, &url)?;
            let mut res = res.text()?;
            loop {
                let params = get_captcha_form(&res);
                res = client
                    .post(BASE_URL)
                    .form(&params)
                    .timeout(Duration::seconds(60).to_std().unwrap())
                    .send()?
                    .text()?;
                if !res.contains("Invalid Captcha") {
                    break;
                }
            }
            let re = Regex::new("href=\"(.*?)\"><img.*?zip").unwrap();
            let mut url = BASE_URL
                .to_string()
                .strip_suffix("/nicgep/app")
                .unwrap()
                .to_string();
            for cap in re.captures_iter(&res) {
                url += &cap[1];
            }
            url = url.replace("amp;", "");
            let res = get_url_response(&client, &url)?;
            let res = res.bytes()?;
            let zipfilepath = folder.join("zipfile.zip");
            let mut file = File::create(&zipfilepath)?;
            file.write_all(res.as_ref())?;
            let file = File::open(&zipfilepath).unwrap();
            let mut archive = zip::ZipArchive::new(file).unwrap();
            let filenames = archive
                .file_names()
                .map(|x| x.to_string())
                .collect::<Vec<String>>();
            let xlfilename = folder.join("excelfile.xls");
            for filename in filenames {
                if filename.contains(".xls") {
                    let mut xlfile = archive.by_name(filename.as_ref()).unwrap();
                    let mut outfile = File::create(&xlfilename).unwrap();
                    io::copy(&mut xlfile, &mut outfile).unwrap();
                    break;
                }
            }
            let mut workbook: Xls<_> = open_workbook(&xlfilename).expect("Cannot open file");
            let mut rowsvec = vec![];
            if let Some(Ok(r)) = workbook.worksheet_range("BoQ1") {
                let mut reqdcellids = vec![-1, -1, -1, -1];
                let mut work_name_para = Paragraph::new();
                let mut header_row = false;
                let mut missing_rates_rows: Vec<(f64, Vec<TableCell>)> = vec![];
                let mut boq_amount = 0.0;
                for row in r.rows() {
                    let mut index: i32 = 0;
                    if reqdcellids[0] == -1 {
                        reqdcellids = vec![-1, -1, -1, -1];
                        for rowi in row {
                            if rowi.is_empty() {
                                break;
                            }
                            let row_string = rowi.to_string();
                            if row_string.contains("Sl.") {
                                reqdcellids[0] = index;
                            } else if row_string.contains("Item Description") {
                                header_row = true;
                                if !sor_opened {
                                    let f_sor = File::open(&sor_path).unwrap();
                                    let sor_reader = BufReader::new(f_sor);
                                    for line_sor in sor_reader.lines() {
                                        let line_str_sor = line_sor.unwrap();
                                        let v_sor: Vec<&str> = line_str_sor.split('|').collect();
                                        sor_map.insert(
                                            v_sor[1].to_string(),
                                            v_sor[2].parse::<f64>().unwrap(),
                                        );
                                        sor_id_map
                                            .insert(v_sor[0].to_string(), v_sor[1].to_string());
                                    }
                                    sor_opened = true;
                                }
                                reqdcellids[1] = index;
                            } else if row_string.contains("Quantity") {
                                reqdcellids[2] = index;
                            } else if row_string.contains("Units") {
                                reqdcellids[3] = index;
                            } else if row_string.contains("Work Name") {
                                work_name_para = create_docx_para(&row_string, false);
                            }
                            index += 1;
                        }
                    }
                    let mut cellsrow = vec![];
                    if reqdcellids[0] != -1 {
                        let desc_str = &row[reqdcellids[1] as usize];
                        if desc_str.is_empty() {
                            break;
                        }
                        for id in &reqdcellids {
                            let cellpara =
                                create_docx_para(row[*id as usize].to_string().as_str(), false);
                            let cell = TableCell::new().add_paragraph(cellpara);
                            cellsrow.push(cell);
                        }
                        if !header_row {
                            let desc_str = desc_str.to_string();
                            if !desc_str.contains(':') {
                                continue;
                            }
                            let qty = &row[reqdcellids[2] as usize].as_f64().unwrap();
                            let desc_str_vec: Vec<&str> = desc_str.split(':').collect();
                            let mut desc_str_processed = desc_str_vec[1].to_string();
                            for desc_piece_ind in 2..desc_str_vec.len() {
                                desc_str_processed += ":";
                                desc_str_processed += desc_str_vec[desc_piece_ind];
                            }
                            desc_str_processed.retain(|c| !c.is_whitespace());
                            desc_str_processed = desc_str_processed
                                .replace(',', "")
                                .to_lowercase()
                                .replace("&amp;", "&")
                                .replace("&#39;", "'")
                                .replace(".", "");
                            if sor_map.contains_key(&desc_str_processed) {
                                let rate = sor_map[&desc_str_processed] * EKM_COST_INDEX;
                                let cellpara = create_docx_para(rate.to_string().as_str(), false);
                                let cell = TableCell::new().add_paragraph(cellpara);
                                cellsrow.push(cell);
                                let ind_amt = qty * rate;
                                let cellpara =
                                    create_docx_para(ind_amt.to_string().as_str(), false);
                                let cell = TableCell::new().add_paragraph(cellpara);
                                cellsrow.push(cell);
                                boq_amount += ind_amt;
                            } else if sor_id_map.contains_key(desc_str_vec[0].trim()) {
                                let rate =
                                    sor_map[sor_id_map[desc_str_vec[0]].as_str()] * EKM_COST_INDEX;
                                let cellpara = create_docx_para(rate.to_string().as_str(), false);
                                let cell = TableCell::new().add_paragraph(cellpara);
                                cellsrow.push(cell);
                                let ind_amt = qty * rate;
                                let cellpara =
                                    create_docx_para(ind_amt.to_string().as_str(), false);
                                let cell = TableCell::new().add_paragraph(cellpara);
                                cellsrow.push(cell);
                                boq_amount += ind_amt;
                            } else {
                                missing_rates_rows.push((*qty, cellsrow.clone()));
                                cellsrow.clear();
                            }
                        } else {
                            let cellpara = create_docx_para("Rate", false);
                            let cell = TableCell::new().add_paragraph(cellpara);
                            cellsrow.push(cell);
                            let cellpara = create_docx_para("Amount", false);
                            let cell = TableCell::new().add_paragraph(cellpara);
                            cellsrow.push(cell);
                        }
                        header_row = false;
                    }
                    if !cellsrow.is_empty() {
                        let row = TableRow::new(cellsrow);
                        rowsvec.push(row);
                    }
                }
                if missing_rates_rows.len() == 1 && amount != -1.0 {
                    if let Some((qty, mut cellsrow)) = missing_rates_rows.pop() {
                        let missing_amount = amount - boq_amount;
                        let rate = missing_amount / qty;
                        let cellpara = create_docx_para(rate.to_string().as_str(), false);
                        let cell = TableCell::new().add_paragraph(cellpara);
                        cellsrow.push(cell);
                        let ind_amt = qty * rate;
                        let cellpara = create_docx_para(ind_amt.to_string().as_str(), false);
                        let cell = TableCell::new().add_paragraph(cellpara);
                        cellsrow.push(cell);
                        missing_rates_rows.push((qty, cellsrow));
                        boq_amount = amount;
                    }
                } else if missing_rates_rows.len() > 1 {
                    //continue;
                }
                if (boq_amount - amount).abs() > 0.001 * amount {
                    send_message(
                        &Op::Debug,
                        format!("Work {}: amount vs schedule mismatch", id),
                        signal_path,
                    );
                }
                while !missing_rates_rows.is_empty() {
                    let row = TableRow::new(missing_rates_rows.pop().unwrap().1);
                    rowsvec.push(row);
                }
                let table_doc = Table::new(rowsvec);
                let mut docxobject = Docx::new();
                docxobject = docxobject.add_paragraph(work_name_para);
                if amount != -1.0 {
                    docxobject = docxobject
                        .add_paragraph(create_docx_para(amount.to_string().as_ref(), false));
                }
                docxobject = docxobject.add_table(table_doc);
                let docxfilename = folder.join(format!("work-{}-date-{}.docx", id, date));
                let file = File::create(&docxfilename).unwrap();
                docxobject.build().pack(file)?;
                if let Ok(()) = send_attachment(op, &docxfilename, signal_path) {
                    succeeded_ids.insert(id);
                }
            }
        }
    }
    let failed: Vec<usize> = (&ids - &succeeded_ids).iter().map(|x| *x).collect();
    let mut message = String::new();
    for i in 0..failed.len() {
        message += failed[i].to_string().as_ref();
        if i == failed.len() - 1 {
            break;
        }
        message += "\n";
    }
    if !newdate {
        f_qrs.set_len(0).unwrap();
        write!(f_qrs, "{}", message)?;
    }
    Ok(())
}

fn get_captcha_form(res: &str) -> HashMap<String, String> {
    let mut params = HashMap::new();
    let re = Regex::new("<input type=\"hidden\" name=\"(.*?)\" value=\"(.*?)\" />").unwrap();
    for cap in re.captures_iter(res) {
        params.insert(cap[1].to_string(), cap[2].to_string());
    }
    params.insert("Submit".to_string(), "Submit".to_string());
    let re = RegexBuilder::new("data:image/png;base64,(.*?)\"")
        .dot_matches_new_line(true)
        .build()
        .unwrap();
    let mut captcha_string = String::new();
    for cap in re.captures_iter(res) {
        captcha_string = extract_captcha_string_from_base64(&cap[1]);
    }
    params.insert("captchaText".to_string(), captcha_string);
    params
}

fn get_url_response(client: &Client, url: &str) -> Result<Response, Error> {
    client
        .get(url)
        .timeout(Duration::seconds(60).to_std().unwrap())
        .send()
}

fn create_docx_para(text: &str, bold: bool) -> Paragraph {
    let mut run = Run::new().add_text(text);
    if bold {
        run = run.bold();
    }
    Paragraph::new().add_run(run)
}

fn extract_captcha_string_from_base64(img_base64: &str) -> String {
    let bytes = base64_light::base64_decode(img_base64);
    let mut dynamic_image = rusty_tesseract::image::load_from_memory(bytes.as_ref()).unwrap();
    let height = dynamic_image.height();
    let width = dynamic_image.width();
    let mut imagebuf = dynamic_image.to_rgba8();
    for i in 0..width {
        for j in 0..height {
            let pixel = dynamic_image.get_pixel(i, j);
            if pixel.channels()[2] == 255 {
                for k in 0..4 {
                    imagebuf.get_pixel_mut(i, j).channels_mut()[k] = 0;
                }
            }
        }
    }
    dynamic_image = DynamicImage::ImageRgba8(imagebuf);
    let imgfd = rusty_tesseract::Image::from_dynamic_image(&dynamic_image).unwrap();
    let default_args = rusty_tesseract::Args {
        psm: Some(7),
        ..Default::default()
    };
    let output = rusty_tesseract::image_to_string(&imgfd, &default_args).unwrap();
    output.trim().to_string()
}

pub fn clear_old_files(num_days: i64) {
    let folder = Path::new(FOLDER_PREFIX);
    let paths = fs::read_dir(folder).unwrap();
    for path in paths {
        let path = path.unwrap().path();
        let filename = path.file_name().unwrap().to_str().unwrap();
        if let Ok(date) = NaiveDate::parse_from_str(filename, "%d-%b-%y") {
            let duration = Local::now().date_naive() - date;
            if duration.num_days() > num_days {
                fs::remove_dir_all(path).unwrap();
            }
        }
    }
}

pub fn get_signal_path() -> String {
    let client = Client::builder()
        .cookie_store(true)
        .user_agent(APP_USER_AGENT)
        .build()
        .unwrap();
    let (download, version) = signal_version_check();
    if download {
        let res = get_url_response(
            &client,
            format!(
                "https://github.com/AsamK/signal-cli/releases/download/v{}/signal-cli-{}.tar.gz",
                &version, &version
            )
            .as_str(),
        );
        let res = res.unwrap().bytes().unwrap();
        std::fs::write("./downloaded_file.tar.gz", &res)
            .expect("Reference proteome download failed for {file_name}");
        let tar_gz = File::open("./downloaded_file.tar.gz").unwrap();
        let tar = flate2::read::GzDecoder::new(tar_gz);
        let mut archive = tar::Archive::new(tar);
        archive.unpack("/opt").unwrap();
        replace_with_correct_libsignal(&version);
    }
    format!("/opt/signal-cli-{}/bin/signal-cli", version)
}

fn signal_version_check() -> (bool, String) {
    let client = Client::builder()
        .cookie_store(true)
        .user_agent(APP_USER_AGENT)
        .build()
        .unwrap();
    let res = get_url_response(
        &client,
        "https://api.github.com/repos/AsamK/signal-cli/releases/latest",
    )
    .unwrap();
    let pq = res.json::<serde_json::Value>().unwrap();
    let q = pq["tag_name"].to_string();
    let p: Vec<&str> = (&q)
        .split(|x| x == '"' || x == 'v')
        .filter(|x| !x.is_empty())
        .collect();
    let mut version_number = String::new();
    let mut new_version_found = false;
    if !p.is_empty() {
        version_number = p[0].to_string();
        let file = File::open("./signal_version_number.txt").unwrap();
        let reader = BufReader::new(file);
        let line = reader.lines().flatten().last().unwrap();
        println!("{}", line);
        if version_number != line {
            println!("Downloading new signal version.");
            fs::write("./signal_version_number.txt", &version_number).unwrap();
            new_version_found = true;
        }
    }
    (new_version_found, version_number)
}

fn get_libsignal_version() -> (bool, String) {
    let client = Client::builder()
        .cookie_store(true)
        .user_agent(APP_USER_AGENT)
        .build()
        .unwrap();
    let res = get_url_response(
        &client,
        "https://api.github.com/repos/exquo/signal-libs-build/releases/latest",
    )
    .unwrap();
    let pq = res.json::<serde_json::Value>().unwrap();
    let q = pq["tag_name"].to_string();
    let q = q.split_once("_").unwrap().1.split_once('"').unwrap().0;
    let mut new_version_found = false;
    let version_number = q.to_string();
    let file = File::open("./libsignal_version_number.txt").unwrap();
    let reader = BufReader::new(file);
    let line = reader.lines().flatten().last().unwrap();
    println!("{}", line);
    if version_number != line {
        println!("Downloading new libsignal version.");
        fs::write("./libsignal_version_number.txt", &version_number).unwrap();
        new_version_found = true;
    }
    (new_version_found, version_number)
}

fn replace_with_correct_libsignal(signal_version: &str) {
    let client = Client::builder()
        .cookie_store(true)
        .user_agent(APP_USER_AGENT)
        .build()
        .unwrap();
    let (download, version) = get_libsignal_version();
    if download
    {
        let url = format!("https://github.com/exquo/signal-libs-build/releases/download/libsignal_{}/libsignal_jni.so-{}-aarch64-unknown-linux-gnu.tar.gz", &version, &version);
        println!("{}", url);
        let res = get_url_response(&client, &url);
        let res = res.unwrap().bytes().unwrap();
        std::fs::write("./downloaded_file.tar.gz", &res)
            .expect("Reference proteome download failed for {file_name}");
        let tar_gz = File::open("./downloaded_file.tar.gz").unwrap();
        let tar = flate2::read::GzDecoder::new(tar_gz);
        let mut archive = tar::Archive::new(tar);
        archive.unpack("/opt").unwrap();
    }
    let paths = fs::read_dir("/opt/signal-cli-".to_string() + signal_version + "/lib").unwrap();
    for path in paths {
        let pathstring = path.unwrap().path();
        let fullname = pathstring.file_name().unwrap().to_str().unwrap();
        if fullname.len() > 16 {
            let newname = &fullname[..16];
            if newname == "libsignal-client" {
                rezip_file(pathstring.to_str().unwrap());
                break;
            }
        }
    }
}

fn rezip_file(libsignal_path: &str) {
    let newlibsignal_path = "/opt/signal-cli-0.13.4/lib/2libsignal-client-0.47.0.jar";
    let file = File::open(&libsignal_path).unwrap();
    let mut newfile = File::create(newlibsignal_path).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let filenames = archive
        .file_names()
        .map(|x| x.to_string())
        .collect::<Vec<String>>();
    let mut zip = zip::ZipWriter::new(&mut newfile);
    let options = zip::write::FileOptions::default();
    for filename in filenames {
        let mut zip_file = archive.by_name(&filename).unwrap();
        let mut writer: Vec<u8> = vec![];
        if filename == "libsignal_jni.so".to_owned() {
            println!("{}", &filename);
            let libspath = "/opt/libsignal_jni.so";
            let mut libsfile = File::open(&libspath).unwrap();
            io::copy(&mut libsfile, &mut writer).unwrap();
        } else {
            io::copy(&mut zip_file, &mut writer).unwrap();
        }
        zip.start_file(filename, options).unwrap();
        zip.write_all(&writer).unwrap();
    }
    zip.finish().unwrap();
    fs::remove_file(libsignal_path).unwrap();
    fs::rename(newlibsignal_path, libsignal_path).unwrap();
}
