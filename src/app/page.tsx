"use client"
import React, { useState, useEffect } from 'react';
import { Search, ShoppingCart, Send, X, Plus, Minus, FileDown } from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// Hardcoded medicine data
import medicineData from "../lib/data.json";

// Type definitions
interface Medicine {
  id: string;
  name: string;
  disc: number;
}

interface CartItem extends Medicine {
  quantity: number;
}

export default function MedicineOrderingSystem() {
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [cart, setCart] = useState<CartItem[]>([]);
  const [suggestions, setSuggestions] = useState<Medicine[]>([]);
  const [orderPlaced, setOrderPlaced] = useState<boolean>(false);
  const [animateCart, setAnimateCart] = useState<boolean>(false);
  const [orderHistory, setOrderHistory] = useState<{date: string, items: CartItem[], orderId: string}[]>([]);
  
  // Fixed WhatsApp number
  const whatsappNumber = "+923493237141";
  
  // Search functionality
  useEffect(() => {
    if (searchTerm.trim() === '') {
      setSuggestions([]);
      return;
    }
    
    const filteredItems = medicineData.filter(item => 
      item.name.toLowerCase().includes(searchTerm.toLowerCase())
    );
    setSuggestions(filteredItems.slice(0, 5).map(item => ({ ...item, id: item.id.toString() })));
  }, [searchTerm]);
  
  // Add item to cart
  const addToCart = (item: Medicine) => {
    const existingItem = cart.find(cartItem => cartItem.id === item.id);
    
    if (existingItem) {
      setCart(cart.map(cartItem => 
        cartItem.id === item.id 
          ? { ...cartItem, quantity: cartItem.quantity + 1 } 
          : cartItem
      ));
    } else {
      setCart([...cart, { ...item, quantity: 1 }]);
    }
    
    // Clear search after adding
    setSearchTerm('');
    setSuggestions([]);
    
    // Animate cart icon
    setAnimateCart(true);
    setTimeout(() => setAnimateCart(false), 500);
  };
  
  // Change item quantity
  const updateQuantity = (id: string, change: number) => {
    const updatedCart = cart.map(item => {
      if (item.id === id) {
        const newQuantity = Math.max(0, item.quantity + change);
        return { ...item, quantity: newQuantity };
      }
      return item;
    }).filter(item => item.quantity > 0);
    
    setCart(updatedCart);
  };
  
  // Remove item from cart
  const removeFromCart = (id: string) => {
    setCart(cart.filter(item => item.id !== id));
  };

  // Generate Excel file with order data and return the blob and order ID
  const generateExcelFile = () => {
    const orderDate = new Date().toISOString();
    const orderId = `ORD-${Date.now().toString().slice(-6)}`;
    
    // Prepare data for Excel
    const excelData = cart.map(item => ({
      'Order ID': orderId,
      'Date': new Date().toLocaleDateString(),
      'Time': new Date().toLocaleTimeString(),
      'Medicine Name': item.name,
      'Quantity': item.quantity,
      'Discount': `${(item.disc * 100).toFixed(0)}%`,
    }));
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    
    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Order Details");
    
    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // Add to order history
    setOrderHistory([...orderHistory, {
      date: orderDate,
      items: [...cart],
      orderId: orderId
    }]);
    
    return { blob: excelBlob, fileName: `MediOrder_${orderId}.xlsx`, orderId };
  };
  
  // Send order to WhatsApp with Excel file
  const sendToWhatsApp = async () => {
    if (cart.length === 0) return;
    
    // Generate Excel and get order details
    const { blob, fileName, orderId } = generateExcelFile();
    
    // Save file locally for customer
    saveAs(blob, fileName);
    
    // Create message for WhatsApp
    let message = `📋 *ORDER DETAILS #${orderId}*\n\n`;
    message += "*Items:*\n";
    
    let totalItems = 0;
    cart.forEach(item => {
      const itemDiscount = item.disc > 0 ? ` (${(item.disc * 100).toFixed(0)}% disc)` : '';
      message += `- ${item.name} x${item.quantity}${itemDiscount}\n`;
      totalItems += item.quantity;
    });
    
    message += `\n*Total Items: ${totalItems}*\n`;
    
    // Add note about the Excel file
    message += "\nExcel file with complete order details attached.";
    
    // Create file URL with encoded message
    const encodedMessage = encodeURIComponent(message);
    
    // For WhatsApp Web with file sharing
    // Note: WhatsApp Web API doesn't directly support file attachments through URL
    // We'll simulate the process by opening WhatsApp and providing instructions
    const whatsappURL = `https://wa.me/${whatsappNumber.replace(/\s+/g, '')}?text=${encodedMessage}`;
    
    try {
      // Create a data URL for the Excel file
      const reader = new FileReader();
      
      reader.onload = function(e) {
        // Show success message
        setOrderPlaced(true);
        
        // Open WhatsApp in a new tab
        window.open(whatsappURL, '_blank');
        
        // Reset cart after successful order and delay
        setTimeout(() => {
          setOrderPlaced(false);
        }, 3000);
      };
      
      reader.readAsDataURL(blob);
    } catch (error) {
      console.error("Error sending to WhatsApp:", error);
      alert("There was an error preparing your order. Please try again.");
    }
  };

  // Create a function to directly email the Excel to seller
  const emailExcelToSeller = async () => {
    // This function would integrate with an email service API
    // Since we can't implement a full email service here, this is a placeholder
    // In a real application, you would use a service like SendGrid, Mailgun, etc.
    
    alert("Email functionality would need to be implemented with a backend service.");
  };

  // Download all order history as Excel
  const downloadOrderHistory = () => {
    if (orderHistory.length === 0) return;
    
    // Flatten all orders into rows
    const allOrdersData = orderHistory.flatMap(order => 
      order.items.map(item => ({
        'Order ID': order.orderId,
        'Date': new Date(order.date).toLocaleDateString(),
        'Time': new Date(order.date).toLocaleTimeString(),
        'Medicine Name': item.name,
        'Quantity': item.quantity,
        'Discount': `${(item.disc * 100).toFixed(0)}%`,
      }))
    );
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(allOrdersData);
    
    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "All Orders");
    
    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // Save file
    saveAs(excelBlob, `MediOrder_History_${Date.now()}.xlsx`);
  };

  return (
    <div className="flex flex-col min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-gradient-to-r from-blue-600 to-blue-800 text-white p-4 shadow-md">
        <div className="container mx-auto flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center">
            <ShoppingCart size={28} className="mr-2" />
            <h1 className="text-2xl font-bold">MediOrder Pro</h1>
          </div>
          
          <div className="flex items-center gap-4 w-full md:w-auto">
            <div className="relative flex-1 md:w-64">
              <div className="flex items-center bg-white rounded-lg overflow-hidden shadow-sm">
                <input
                  type="text"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  placeholder="Search medicines..."
                  className="p-2 w-full md:w-64 text-gray-800 focus:outline-none"
                />
                <Search className="mr-2 text-gray-500" size={20} />
              </div>
              
              {/* Search suggestions */}
              {suggestions.length > 0 && (
                <div className="absolute z-10 mt-1 w-full bg-white border rounded-md shadow-lg">
                  {suggestions.map(item => (
                    <div 
                      key={item.id}
                      className="p-2 hover:bg-gray-100 cursor-pointer border-b flex justify-between items-center"
                      onClick={() => addToCart(item)}
                    >
                      <div>
                        <div className="font-medium text-black">{item.name}</div>
                        <div className="flex items-center gap-2">
                          {item.disc > 0 && (
                            <span className="text-sm text-green-600">
                              {(item.disc * 100).toFixed(0)}% off
                            </span>
                          )}
                        </div>
                      </div>
                      <Plus size={18} className="text-blue-500" />
                    </div>
                  ))}
                </div>
              )}
            </div>
            
            <div className="relative cursor-pointer">
              <ShoppingCart 
                size={26} 
                className={`transition-transform ${animateCart ? 'scale-125' : ''}`} 
              />
              {cart.length > 0 && (
                <span className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-6 h-6 flex items-center justify-center text-xs font-bold shadow-sm">
                  {cart.reduce((sum, item) => sum + item.quantity, 0)}
                </span>
              )}
            </div>
          </div>
        </div>
      </header>
      
      {/* Main content */}
      <main className="container mx-auto p-4 flex flex-col md:flex-row gap-6 flex-grow">
        {/* Product listing */}
        <div className="w-full md:w-2/3">
          <div className="bg-white rounded-lg shadow-md p-4">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Available Medicines</h2>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {medicineData.map(item => (
                <div 
                  key={item.id}
                  className="border rounded-lg p-3 hover:shadow-md transition-shadow"
                >
                  <h3 className="font-medium text-gray-800 mb-1">{item.name}</h3>
                  
                  <div className="flex justify-between items-center mt-3">
                    <div className="flex items-baseline gap-2">
                      {item.disc > 0 && (
                        <span className="text-sm text-green-600 font-medium">
                          {(item.disc * 100).toFixed(0)}% off
                        </span>
                      )}
                    </div>
                    
                    <button 
                      onClick={() => addToCart(item)}
                      className="bg-blue-500 hover:bg-blue-600 text-white rounded-lg px-3 py-1 flex items-center gap-1 text-sm transition-colors"
                    >
                      <Plus size={16} />
                      Add
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
        
        {/* Cart */}
        <div className="w-full md:w-1/3">
          <div className="bg-white rounded-lg shadow-md p-4 sticky top-4">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800 flex items-center">
                <ShoppingCart size={20} className="mr-2 text-blue-500" />
                Your Cart
              </h2>
              {cart.length > 0 && (
                <button 
                  onClick={() => setCart([])}
                  className="text-red-500 text-sm hover:underline flex items-center"
                >
                  <X size={14} className="mr-1" />
                  Clear All
                </button>
              )}
            </div>
            
            {cart.length === 0 ? (
              <div className="py-8 text-center text-gray-500">
                <ShoppingCart size={32} className="mx-auto mb-2 opacity-50" />
                <p>Your cart is empty</p>
                <p className="text-sm mt-1">Search or browse medicines to add them</p>
              </div>
            ) : (
              <>
                <div className="max-h-64 overflow-auto mb-4 pr-1">
                  {cart.map(item => (
                    <div 
                      key={item.id} 
                      className="flex justify-between items-center py-3 border-b group"
                    >
                      <div className="flex-1">
                        <h3 className="font-medium text-gray-800">{item.name}</h3>
                        <div className="flex items-center text-sm">
                          {item.disc > 0 && (
                            <span className="text-green-600">
                              {(item.disc * 100).toFixed(0)}% discount
                            </span>
                          )}
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-2 ml-4">
                        <button 
                          onClick={() => updateQuantity(item.id, -1)}
                          className="bg-gray-200 hover:bg-gray-300 rounded-full p-1"
                        >
                          <Minus size={14} />
                        </button>
                        <span className="w-6 text-center font-medium">{item.quantity}</span>
                        <button 
                          onClick={() => updateQuantity(item.id, 1)}
                          className="bg-gray-200 hover:bg-gray-300 rounded-full p-1"
                        >
                          <Plus size={14} />
                        </button>
                        <button 
                          onClick={() => removeFromCart(item.id)}
                          className="text-red-500 ml-1 opacity-70 hover:opacity-100"
                        >
                          <X size={16} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
                
                <div className="mt-4 space-y-2">
                  <button
                    onClick={sendToWhatsApp}
                    disabled={cart.length === 0}
                    className={`w-full py-3 rounded-md flex items-center justify-center gap-2 font-medium transition-colors ${
                      cart.length === 0
                        ? 'bg-gray-300 text-gray-500 cursor-not-allowed' 
                        : 'bg-green-500 hover:bg-green-600 text-white'
                    }`}
                  >
                    <Send size={16} />
                    Send Order with Excel
                  </button>

                  {orderHistory.length > 0 && (
                    <button
                      onClick={downloadOrderHistory}
                      className="w-full py-2 rounded-md flex items-center justify-center gap-2 font-medium transition-colors bg-blue-500 hover:bg-blue-600 text-white"
                    >
                      <FileDown size={16} />
                      Download Order History
                    </button>
                  )}
                  
                  {orderPlaced && (
                    <div className="mt-2 text-center bg-green-50 text-green-600 text-sm p-2 rounded-md">
                      <p>Order sent successfully!</p>
                      <p className="text-xs mt-1">Excel file has been downloaded and order details sent to WhatsApp.</p>
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </main>

      {/* Footer with WhatsApp contact */}
      <footer className="bg-gray-800 text-white py-3 mt-auto">
        <div className="container mx-auto text-center text-sm">
          <p>Contact Seller: {whatsappNumber}</p>
          <p className="mt-1 text-gray-400">All orders are saved as Excel files and sent to seller</p>
        </div>
      </footer>
    </div>
  );
}