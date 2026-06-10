const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, LevelFormat, HeadingLevel
} = require('docx');
const fs = require('fs');

// ── font / spacing helpers (matching original report) ────────────────────────
const F = "Times New Roman";

function kbTitle(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120, line: 312, lineRule: "auto" },
    indent: { firstLine: 567 },
    alignment: AlignmentType.BOTH,
    children: [new TextRun({ text, bold: true, size: 26, font: F })]
  });
}

function bullet(label, content) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 120, after: 120, line: 312, lineRule: "auto" },
    alignment: AlignmentType.BOTH,
    children: [
      new TextRun({ text: label, bold: true, size: 26, font: F }),
      new TextRun({ text: " " + content, size: 26, font: F })
    ]
  });
}

function spacer() {
  return new Paragraph({
    children: [new TextRun("")],
    spacing: { before: 60, after: 60 }
  });
}

// ── 5 kịch bản ───────────────────────────────────────────────────────────────
const scenarios = [
  {
    title: "Kịch bản 1: Xây dựng hạ tầng mạng nền tảng (VLAN, Định tuyến & DHCP)",
    mo_ta:
      "Thiết lập toàn bộ nền tảng hạ tầng mạng cho công ty: phân chia các phòng ban " +
      "thành các vùng mạng riêng biệt (VLAN), cấu hình định tuyến liên VLAN thông qua " +
      "L3-Switch và triển khai cấp phát IP tự động (DHCP) cho từng phòng ban. Đây là " +
      "kịch bản nền tảng, tất cả các kịch bản tiếp theo đều được xây dựng trên cơ sở này.",
    cach_thuc_hien:
      "Khởi tạo 5 VLAN tương ứng với 5 phòng ban: VLAN 10 – Giám Đốc (192.168.10.0/24), " +
      "VLAN 20 – Kế Toán (192.168.20.0/24), VLAN 30 – Truyền Thông (192.168.30.0/24), " +
      "VLAN 40 – Nhân Sự (192.168.40.0/24), VLAN 50 – IT (192.168.50.0/24). " +
      "Cấu hình cổng Trunk (802.1Q) giữa L3-Switch và các Access Switch. " +
      "Tạo SVI (Switch Virtual Interface) trên L3-Switch1 và L3-Switch2 làm Default Gateway cho từng VLAN. " +
      "Khởi tạo DHCP Pool riêng biệt cho mỗi VLAN, loại trừ các IP tĩnh dành cho " +
      "Gateway, Firewall, Server và thiết bị phòng IT. " +
      "Đặt IP tĩnh cho các thiết bị trong phòng IT (PC5_IT: 192.168.50.10, PC_IT: 192.168.50.11, " +
      "Server0_IT: 192.168.50.100, Server1_IT: 192.168.50.101). " +
      "Cấu hình Default Route trên L3-Switch (ip route 0.0.0.0 0.0.0.0 192.168.200.1) " +
      "trỏ về Firewall để các VLAN có thể ra ngoài Internet.",
    ket_qua:
      "Các thiết bị đầu cuối nhận đúng dải IP, Subnet Mask và Default Gateway của phòng ban mình. " +
      "Các VLAN định tuyến được với nhau thông qua L3-Switch (ping thành công giữa các phòng ban " +
      "khi chưa áp dụng ACL). Thiết bị phòng IT nhận IP tĩnh đúng như quy hoạch."
  },
  {
    title: "Kịch bản 2: Đảm bảo tính sẵn sàng cao và bảo mật lớp chuyển mạch",
    mo_ta:
      "Kịch bản này giải quyết 2 mối đe dọa: (1) Sập toàn bộ hệ thống khi một Core Switch " +
      "gặp sự cố (Single Point of Failure) – xử lý bằng HSRP; (2) Các tấn công vật lý từ " +
      "bên trong như cắm thiết bị lạ, giả mạo DHCP Server, phá hoại cấu trúc Spanning Tree – " +
      "xử lý bằng Layer 2 Security.",
    cach_thuc_hien:
      "Phần HSRP: Cấu hình HSRP trên L3-Switch1 (Active) và L3-Switch2 (Standby) cho từng VLAN. " +
      "Khởi tạo Virtual IP làm Default Gateway dùng chung " +
      "(VLAN10: 192.168.10.254, VLAN20: 192.168.20.254, VLAN30: 192.168.30.254, " +
      "VLAN40: 192.168.40.254, VLAN50: 192.168.50.254). " +
      "Phần Layer 2 Security: Bật Port Security trên tất cả cổng Access của " +
      "SW-GIANDOC, SW-KETOAN, SW-TRUYENTHONG, SW-NHANSU, SW-IT – giới hạn 1 MAC/cổng, " +
      "vi phạm thì chuyển sang err-disable. " +
      "Bật DHCP Snooping toàn Switch để chặn DHCP Server giả mạo. " +
      "Kích hoạt BPDU Guard trên tất cả cổng Access để chống cắm Switch lạ.",
    ket_qua:
      "Khi tắt đột ngột L3-Switch1, L3-Switch2 tự động lên Active trong vòng vài giây, " +
      "các máy trạm không bị gián đoạn kết nối. " +
      "Nếu nhân viên cắm bộ phát Wifi cá nhân vào cổng mạng, cổng đó lập tức chuyển sang " +
      "err-disable và ghi log cảnh báo. DHCP giả mạo bị chặn hoàn toàn."
  },
  {
    title: "Kịch bản 3: Thiết lập Firewall bảo vệ vùng DMZ và kết nối Internet (NAT)",
    mo_ta:
      "Xây dựng chốt chặn an ninh trung tâm trên Firewall ASA 5506-X với 3 nhiệm vụ: " +
      "(1) Bảo vệ vùng DMZ chứa WEB-SERVER (192.168.150.10), MAL-SERVER (192.168.150.11), " +
      "FTP-SERVER (192.168.150.12); " +
      "(2) Kiểm soát toàn bộ luồng dữ liệu giữa Inside – DMZ – Outside theo Security Level; " +
      "(3) Cung cấp kết nối Internet cho nội bộ và xuất bản dịch vụ ra ngoài qua NAT.",
    cach_thuc_hien:
      "Khai báo 3 interface trên Firewall: Outside (Gi1/1 – 192.168.100.2/30, Security Level 0), " +
      "Inside (Gi1/2 – 192.168.200.1/30, Security Level 100), " +
      "DMZ (Gi1/3 – 192.168.150.1/24, Security Level 50). " +
      "Cấu hình Default Route trên Firewall trỏ ra Router4 " +
      "(route outside 0.0.0.0 0.0.0.0 192.168.100.1). " +
      "Cấu hình Static Route về các VLAN nội bộ qua L3-Switch1 " +
      "(route inside 192.168.0.0 255.255.0.0 192.168.200.2). " +
      "Thêm Static Route trên Router4 về mạng nội bộ " +
      "(ip route 192.168.0.0 255.255.0.0 203.0.113.2). " +
      "Bật inspect icmp trong policy-map để Firewall cho phép ICMP reply đi qua. " +
      "Cấu hình Dynamic NAT/PAT ẩn toàn bộ IP Private khi ra Internet. " +
      "Cấu hình Static NAT ánh xạ WEB-SERVER và FTP-SERVER ra IP Public.",
    ket_qua:
      "Các PC nội bộ duyệt web và ping ra Internet thành công. " +
      "Người dùng Internet truy cập được WEB-SERVER và FTP-SERVER của công ty. " +
      "Mạng DMZ được cô lập – máy chủ trong DMZ không thể tự kết nối vào mạng Inside. " +
      "Firewall ghi nhận và kiểm soát toàn bộ luồng dữ liệu qua lại giữa 3 vùng."
  },
  {
    title: "Kịch bản 4: Kiểm soát truy cập nội bộ (ACL) và phân quyền phòng ban",
    mo_ta:
      "Thực thi nguyên tắc \"đặc quyền tối thiểu\" – mỗi phòng ban chỉ được truy cập đúng " +
      "những tài nguyên cần thiết. Kịch bản này bao gồm 2 chính sách cụ thể đã được kiểm thử: " +
      "(1) Phòng Giám Đốc có thể ping đến tất cả các phòng nhưng các phòng khác không thể " +
      "ping ngược lại; (2) Chỉ phòng IT mới được truy cập vào Server0_IT và Server1_IT.",
    cach_thuc_hien:
      "Tạo Extended ACL \"BLOCK_PING_TO_GD\" trên L3-Switch1: " +
      "Cho phép ICMP từ VLAN 10 đi bất kỳ đâu (permit icmp 192.168.10.0 0.0.0.255 any). " +
      "Chặn ICMP echo-request từ các VLAN 20/30/40/50 vào VLAN 10 " +
      "(deny icmp [subnet] 0.0.0.255 192.168.10.0 0.0.0.255 echo). " +
      "Áp ACL này inbound lên SVI VLAN 20, 30, 40, 50 – " +
      "lưu ý phải dùng từ khóa \"echo\" để chỉ chặn ping đi, không chặn reply về. " +
      "Tạo Extended ACL \"PROTECT_IT_SERVER\" chặn truy cập TCP/UDP từ VLAN 10/20/30/40 " +
      "vào dải 192.168.50.100–101, chỉ cho phép VLAN 50 (IT) truy cập. " +
      "Áp ACL tương tự lên L3-Switch2 để đảm bảo chính sách không bị bypass qua Switch dự phòng.",
    ket_qua:
      "PC0_GD (192.168.10.x) ping thành công đến PC1_KT, PC3_TT, PC4_NS, PC5_IT. " +
      "PC1_KT, PC3_TT, PC4_NS, PC5_IT ping đến PC0_GD đều nhận Request Timeout. " +
      "PC5_IT và PC_IT SSH/FTP vào Server0_IT và Server1_IT thành công. " +
      "PC các phòng ban khác bị từ chối kết nối vào vùng Server IT."
  },
  {
    title: "Kịch bản 5: Thiết lập kết nối từ xa an toàn (Remote Access VPN)",
    mo_ta:
      "Cung cấp đường truyền được mã hóa cho nhân viên làm việc ngoài văn phòng có thể " +
      "truy cập an toàn vào tài nguyên nội bộ (Server0_IT, Server1_IT) từ bất kỳ đâu " +
      "thông qua mạng Internet công cộng, đảm bảo dữ liệu không bị đánh cắp trên đường truyền.",
    cach_thuc_hien:
      "Khởi tạo chính sách mã hóa IKE Phase 1 trên Firewall ASA 5506-X " +
      "(thuật toán: AES-256, SHA-1, DH Group 2). " +
      "Cấu hình IPsec Transform-Set cho IKE Phase 2 " +
      "(ESP-AES-256, ESP-SHA-HMAC). " +
      "Tạo IP Pool cấp phát địa chỉ ảo cho VPN Client (ví dụ: 192.168.60.0/24). " +
      "Tạo tài khoản xác thực cho người dùng từ xa (username/password cục bộ trên ASA). " +
      "Cấu hình Group Policy và Tunnel Group cho Remote Access VPN. " +
      "Mở ACL trên cổng Outside cho phép giao thức IKE (UDP 500) và IPsec (UDP 4500) đi vào.",
    ket_qua:
      "Nhân viên sử dụng Cisco VPN Client kết nối từ mạng ngoài, xác thực thành công " +
      "và nhận địa chỉ IP ảo trong dải 192.168.60.0/24. " +
      "Sau khi kết nối VPN, nhân viên truy cập được Server0_IT (192.168.50.100) " +
      "và Server1_IT (192.168.50.101) như đang ở trong văn phòng. " +
      "Toàn bộ lưu lượng được mã hóa – không thể bị nghe lén trên đường truyền Internet."
  }
];

// ── Build document ────────────────────────────────────────────────────────────
const children = [];

// Section title
children.push(new Paragraph({
  spacing: { before: 200, after: 200, line: 312, lineRule: "auto" },
  alignment: AlignmentType.BOTH,
  children: [new TextRun({
    text: "2.4.5. Các kịch bản triển khai hệ thống mạng",
    bold: true, size: 26, font: F
  })]
}));

scenarios.forEach((s, i) => {
  children.push(kbTitle(s.title));
  children.push(bullet("Mô tả:", s.mo_ta));
  children.push(bullet("Cách thực hiện:", s.cach_thuc_hien));
  children.push(bullet("Kết quả mong đợi:", s.ket_qua));
  if (i < scenarios.length - 1) children.push(spacer());
});

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "\u2013",
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: {
            indent: { left: 1560, hanging: 360 }
          },
          run: { font: F }
        }
      }]
    }]
  },
  styles: {
    default: {
      document: { run: { font: F, size: 26 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/5_Kich_Ban_Nhom12.docx', buf);
  console.log('Done');
});
