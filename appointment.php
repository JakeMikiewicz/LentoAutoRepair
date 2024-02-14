<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Retrieve form data
    $name = $_POST['name'];
    $email = $_POST['email'];
    $date = $_POST['date'];
    $time = $_POST['time'];

    // Connect to the Access database (change the path to your .mdb or .accdb file)
    $dbPath = 'path/to/your/access-database.accdb';
    $connection = new COM("ADODB.Connection");
    $connectionString = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=$dbPath";
    $connection->open($connectionString);

    // Prepare and execute SQL query to insert appointment
    $query = "INSERT INTO Appointments (Name, Email, AppointmentDate, AppointmentTime) VALUES ('$name', '$email', '$date', '$time')";
    $connection->Execute($query);

    // Close the database connection
    $connection->close();
}
?>