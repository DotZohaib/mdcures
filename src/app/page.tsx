/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-empty-interface */
/* eslint-disable @typescript-eslint/no-unused-vars */

"use client"
import React, { useState, useEffect } from 'react';
import { Search, ShoppingCart, Send, X, Plus, Minus, FileDown, AlertCircle, Check, ShoppingBag, Bell } from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import Link from 'next/link';

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

interface Toast {
  id: string;
  message: string;
  type: 'success' | 'error' | 'info';
}

export default function MedicineOrderingSystem() {
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [cart, setCart] = useState<CartItem[]>([]);
  const [suggestions, setSuggestions] = useState<Medicine[]>([]);
  const [orderPlaced, setOrderPlaced] = useState<boolean>(false);
  const [animateCart, setAnimateCart] = useState<boolean>(false);
  const [orderHistory, setOrderHistory] = useState<{date: string, items: CartItem[], orderId: string}[]>([]);
  const [isCartPopupOpen, setIsCartPopupOpen] = useState<boolean>(false);
  const [toasts, setToasts] = useState<Toast[]>([]);
  
  // Fixed WhatsApp number
  const whatsappNumber = "+923403004201";
  
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
  
  // Toast notification system
  const addToast = (message: string, type: 'success' | 'error' | 'info') => {
    const id = Date.now().toString();
    setToasts(prev => [...prev, { id, message, type }]);
    
    // Auto-remove toast after 3 seconds
    setTimeout(() => {
      setToasts(prev => prev.filter(toast => toast.id !== id));
    }, 3000);
  };
  
  // Close cart popup when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as HTMLElement;
      if (isCartPopupOpen && !target.closest('#cart-popup') && !target.closest('#cart-trigger')) {
        setIsCartPopupOpen(false);
      }
    };
    
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isCartPopupOpen]);
  
  // Add item to cart
  const addToCart = (item: Medicine) => {
    const existingItem = cart.find(cartItem => cartItem.id === item.id);
    
    if (existingItem) {
      setCart(cart.map(cartItem => 
        cartItem.id === item.id 
          ? { ...cartItem, quantity: cartItem.quantity + 1 } 
          : cartItem
      ));
      addToast(`Increased quantity of ${item.name}`, 'success');
    } else {
      setCart([...cart, { ...item, quantity: 1 }]);
      addToast(`Added ${item.name} to cart`, 'success');
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
    const item = cart.find(item => item.id === id);
    if (!item) return;
    
    const updatedCart = cart.map(item => {
      if (item.id === id) {
        const newQuantity = Math.max(0, item.quantity + change);
        return { ...item, quantity: newQuantity };
      }
      return item;
    }).filter(item => item.quantity > 0);
    
    setCart(updatedCart);
    
    if (change > 0) {
      addToast(`Increased quantity of ${item.name}`, 'info');
    } else if (change < 0) {
      if (item.quantity === 1) {
        addToast(`Removed ${item.name} from cart`, 'info');
      } else {
        addToast(`Decreased quantity of ${item.name}`, 'info');
      }
    }
  };
  
  // Remove item from cart
  const removeFromCart = (id: string) => {
    const item = cart.find(item => item.id === id);
    if (item) {
      setCart(cart.filter(item => item.id !== id));
      addToast(`Removed ${item.name} from cart`, 'info');
    }
  };
  
  // Clear cart
  const clearCart = () => {
    if (cart.length > 0) {
      setCart([]);
      addToast('Cart cleared', 'info');
    }
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
    if (cart.length === 0) {
      addToast('Your cart is empty', 'error');
      return;
    }
    
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
    message += "\n Your order has been placed. ";
    
    // Create file URL with encoded message
    const encodedMessage = encodeURIComponent(message);
    
    // For WhatsApp Web with file sharing
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
          setCart([]);
          setOrderPlaced(false);
          addToast('Order placed successfully! WhatsApp opened with details.', 'success');
        }, 1000);
      };
      
      reader.readAsDataURL(blob);
    } catch (error) {
      console.error("Error sending to WhatsApp:", error);
      addToast('Error sending order. Please try again.', 'error');
    }
  };

  // Download all order history as Excel
  const downloadOrderHistory = () => {
    if (orderHistory.length === 0) {
      addToast('No order history available', 'error');
      return;
    }
    
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
    addToast('Order history downloaded successfully', 'success');
  };

  // Get total number of items in cart
  const totalItemsInCart = cart.reduce((sum, item) => sum + item.quantity, 0);

  return (
    <div className="flex flex-col min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-gradient-to-r from-blue-600 to-blue-800 text-white p-4 shadow-md">
        <div className="container mx-auto flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center">
            <ShoppingBag size={28} className="mr-2" />
            <h1 className="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-green-400 to-blue-500 drop-shadow-lg">
              <b className="text-white">MD</b> <span className="italic text-gray-300">Cure</span>
            </h1>
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
            
            {/* Cart trigger */}
            <div 
              id="cart-trigger"
              className="relative cursor-pointer"
              onClick={() => setIsCartPopupOpen(!isCartPopupOpen)}
            >
              <ShoppingCart 
                size={26} 
                className={`transition-transform ${animateCart ? 'scale-125' : ''}`} 
              />
              {cart.length > 0 && (
                <span className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-6 h-6 flex items-center justify-center text-xs font-bold shadow-sm">
                  {totalItemsInCart}
                </span>
              )}
            </div>
          </div>
        </div>
      </header>
      
      {/* Cart Popup */}
      {isCartPopupOpen && (
        <div
          id="cart-popup"
          className="fixed right-4 top-20 w-80 bg-white z-50 rounded-lg shadow-lg overflow-hidden transition-all transform origin-top-right"
        >
          <div className="p-3 bg-gradient-to-r from-blue-600 to-blue-800 text-white flex justify-between items-center">
            <div className="flex items-center">
              <ShoppingCart size={18} className="mr-2" />
              <span className="font-bold">Your Cart</span>
            </div>
            <div className="flex gap-2">
              {cart.length > 0 && (
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    clearCart();
                  }}
                  className="text-white text-xs hover:underline flex items-center"
                >
                  <X size={14} className="mr-1" />
                  Clear
                </button>
              )}
              <button 
                onClick={(e) => {
                  e.stopPropagation();
                  setIsCartPopupOpen(false);
                }}
                className="text-white hover:text-gray-200"
              >
                <X size={18} />
              </button>
            </div>
          </div>
          
          {cart.length === 0 ? (
            <div className="p-4 text-center text-gray-500">
              <ShoppingCart size={32} className="mx-auto mb-2 opacity-50" />
              <p>Your cart is empty</p>
              <p className="text-sm mt-1">Search or browse medicines to add them</p>
            </div>
          ) : (
            <>
              <div className="max-h-64 overflow-auto p-2">
                {cart.map(item => (
                  <div 
                    key={item.id} 
                    className="flex justify-between items-center py-2 px-1 border-b hover:bg-gray-50 group"
                  >
                    <div className="flex-1">
                      <h3 className="font-medium text-gray-800 text-sm">{item.name}</h3>
                      {item.disc > 0 && (
                        <span className="text-xs text-green-600">
                          {(item.disc * 100).toFixed(0)}% discount
                        </span>
                      )}
                    </div>
                    
                    <div className="flex items-center gap-1 ml-2">
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          updateQuantity(item.id, -1);
                        }}
                        className="bg-gray-200 hover:bg-gray-300 rounded-full p-1"
                      >
                        <Minus size={12} />
                      </button>
                      <span className="w-5 text-center text-sm font-medium">{item.quantity}</span>
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          updateQuantity(item.id, 1);
                        }}
                        className="bg-gray-200 hover:bg-gray-300 rounded-full p-1"
                      >
                        <Plus size={12} />
                      </button>
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          removeFromCart(item.id);
                        }}
                        className="text-red-500 ml-1 opacity-70 hover:opacity-100"
                      >
                        <X size={14} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
              
              <div className="p-3 bg-gray-50 border-t">
                <div className="flex justify-between text-sm mb-2">
                  <span className="font-medium">Total Items:</span>
                  <span className="font-bold">{totalItemsInCart}</span>
                </div>
                
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    sendToWhatsApp();
                  }}
                  className="w-full py-2 rounded-md bg-green-500 hover:bg-green-600 text-white font-medium flex items-center justify-center gap-2 text-sm"
                >
                  <Send size={14} />
                  Send Order 
                </button>
                
                <Link href="#cart">
                  <button
                    onClick={() => setIsCartPopupOpen(false)}
                    className="w-full mt-2 py-1.5 rounded-md bg-blue-500 hover:bg-blue-600 text-white text-xs flex items-center justify-center gap-1"
                  >
                    <ShoppingCart size={12} />
                    View Full Cart
                  </button>
                </Link>
              </div>
            </>
          )}
        </div>
      )}
      
      {/* Main content */}
      <main className="container mx-auto p-4 flex flex-col md:flex-row gap-6 flex-grow">
        {/* Product listing */}
        <div className="w-full md:w-2/3">
          <div className="bg-white rounded-lg shadow-md p-4">
            <h2 className="text-xl font-bold text-gray-800 mb-4 flex items-center">
              <Bell size={20} className="mr-2 text-blue-500" />
              Available Medicines
            </h2>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {medicineData.map(item => (
                <div 
                  key={item.id}
                  className="border rounded-lg p-3 hover:shadow-md transition-shadow bg-white hover:bg-blue-50"
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
                      onClick={() => addToCart(item as unknown as Medicine)}
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
        <div id="cart" className="w-full md:w-1/3">
          <div className="bg-white rounded-lg shadow-md p-4 sticky top-4">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800 flex items-center">
                <ShoppingCart size={20} className="mr-2 text-blue-500" />
                Your Cart
              </h2>
              {cart.length > 0 && (
                <button 
                  onClick={clearCart}
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
                      className="flex justify-between items-center py-3 border-b group hover:bg-gray-50 rounded px-2"
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
                  <div className="bg-blue-50 p-3 rounded-md text-sm mb-3">
                    <div className="flex justify-between mb-1">
                      <span>Total Items:</span>
                      <span className="font-bold">{totalItemsInCart}</span>
                    </div>
                  </div>
                  
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
                    Send Order
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

      {/* Toast Notifications */}
      <div className="fixed bottom-4 right-4 z-50 flex flex-col gap-2 max-w-sm">
        {toasts.map(toast => (
          <div 
            key={toast.id} 
            className={`
              flex items-center p-3 rounded-lg shadow-lg text-white transform transition-all duration-300 animate-slideIn
              ${toast.type === 'success' ? 'bg-green-500' : ''}
              ${toast.type === 'error' ? 'bg-red-500' : ''}
              ${toast.type === 'info' ? 'bg-blue-500' : ''}
            `}
            style={{ animation: 'slideIn 0.3s ease-out, fadeOut 0.3s ease-in 2.7s' }}
          >
            <div className="mr-2">
              {toast.type === 'success' && <Check size={18} />}
              {toast.type === 'error' && <AlertCircle size={18} />}
              {toast.type === 'info' && <Bell size={18} />}
            </div>
            <p className="text-sm">{toast.message}</p>
            <button 
              onClick={() => setToasts(toasts.filter(t => t.id !== toast.id))}
              className="ml-auto text-white hover:text-gray-200"
            >
              <X size={16} />
            </button>
          </div>
        ))}
      </div>

      {/* Footer with WhatsApp contact */}
      <footer className="bg-gradient-to-r from-blue-600 to-blue-800 text-white py-3 mt-auto">
  <div className="container mx-auto text-center text-sm space-y-2">
    <p className="mt-1 text-gray-200">&copy; MBCure. All rights reserved.</p>
    <p className="text-gray-200">
      Contact: <strong>Babar Ali Dayo</strong> –&nbsp;
      <a
        href="tel:+923403004201"
        className="underline hover:text-gray-100"
      >
        +92 340 300 4201
      </a>
      &nbsp;|&nbsp;
      <a
        href="https://wa.me/923403004201"
        target="_blank"
        rel="noopener noreferrer"
        className="inline-flex items-center underline hover:text-gray-100"
      >
        WhatsApp Chat
      </a>
    </p>
  </div>
</footer>

      {/* CSS Animation for Toast */}
      <style jsx global>{`
        @keyframes slideIn {
          0% { transform: translateX(100%); opacity: 0; }
          100% { transform: translateX(0); opacity: 1; }
        }
        
        @keyframes fadeOut {
          0% { opacity: 1; }
          100% { opacity: 0; }
        }
        
        .animate-slideIn {
          animation: slideIn 0.3s ease-out;
        }
      `}</style>
    </div>
  );
}
